"""Index GitHub Issues into a Microsoft 365 Copilot connector (Microsoft Graph external connection).

Commands:
    python github_issues_connector.py fetch    # list issues from GitHub only (no Microsoft Graph calls)
    python github_issues_connector.py setup    # create the external connection and register the schema
    python github_issues_connector.py ingest   # push all open and closed issues into the connection
    python github_issues_connector.py all      # setup + ingest
    python github_issues_connector.py delete   # delete the connection (and its indexed issues)
"""

import argparse
import os
import re
import sys
import time

import msal
import requests
from dotenv import load_dotenv

GRAPH = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GITHUB_API = "https://api.github.com"
RETRY_STATUSES = (429, 503, 504)
MAX_RETRIES = 5
REPO_RE = re.compile(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$")
CONNECTION_ID_RE = re.compile(r"^[A-Za-z0-9]{3,32}$")

# Flat list of properties Microsoft 365 Copilot can search, filter and show for each issue.
SCHEMA = {
    "baseType": "microsoft.graph.externalItem",
    "properties": [
        {
            "name": "title",
            "type": "String",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": True,
            "labels": ["title"],
        },
        {
            "name": "status",
            "type": "String",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {"name": "url", "type": "String", "isRetrievable": True, "labels": ["url"]},
        {"name": "issueNumber", "type": "Int64", "isRetrievable": True, "isQueryable": True},
        {"name": "author", "type": "String", "isRetrievable": True, "isQueryable": True},
        {
            "name": "labels",
            "type": "StringCollection",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": True,
        },
        {
            "name": "updatedAt",
            "type": "DateTime",
            "isRetrievable": True,
            "isQueryable": True,
            "labels": ["lastModifiedDateTime"],
        },
    ],
}


class ConnectorError(Exception):
    pass


def env(name: str, default: str = "") -> str:
    return (os.getenv(name) or default).strip()


def require_env(*names: str):
    missing = [n for n in names if not env(n)]
    if missing:
        raise ConnectorError(f"Missing environment variable(s): {', '.join(missing)}. See .env.example.")


def github_repo() -> str:
    require_env("GITHUB_REPO")
    repo = env("GITHUB_REPO")
    if not REPO_RE.match(repo):
        raise ConnectorError(f"GITHUB_REPO must be 'owner/repo', got '{repo}'")
    return repo


def connection_id() -> str:
    value = env("CONNECTION_ID", "githubissuespoc")
    if not CONNECTION_ID_RE.match(value):
        raise ConnectorError("CONNECTION_ID must be 3-32 alphanumeric characters.")
    if value.lower().startswith("microsoft"):
        raise ConnectorError("CONNECTION_ID must not start with 'Microsoft'.")
    return value


def ensure_ok(resp: requests.Response, action: str):
    if resp.ok:
        return
    hint = ""
    if resp.status_code in (401, 403):
        hint = (
            "\nHint: the app needs the Microsoft Graph *application* permissions "
            "ExternalConnection.ReadWrite.OwnedBy and ExternalItem.ReadWrite.OwnedBy, "
            "with admin consent granted in the tenant."
        )
    elif resp.status_code == 412:
        hint = (
            "\nHint: the tenant is not ready for Copilot connectors. It needs an active "
            "Microsoft 365 subscription; new tenants must wait about an hour before the first connection."
        )
    raise ConnectorError(f"{action} failed: HTTP {resp.status_code}\n{resp.text}{hint}")


# --- GitHub ---------------------------------------------------------------------------------

def fetch_issues(repo: str) -> tuple[list[dict], int]:
    """Return all open and closed issues for the repo, excluding pull requests, and the skipped PR count."""
    session = requests.Session()
    session.headers.update({"Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"})
    token = env("GITHUB_TOKEN")
    if token:
        session.headers["Authorization"] = f"Bearer {token}"

    url = f"{GITHUB_API}/repos/{repo}/issues"
    params = {"state": "all", "per_page": 100}
    issues, skipped_prs = [], 0
    while url:
        resp = session.get(url, params=params, timeout=60)
        params = None  # the "next" link already carries the query string
        if not resp.ok:
            message = resp.json().get("message", resp.text) if resp.content else resp.text
            raise ConnectorError(f"GitHub request failed: HTTP {resp.status_code} - {message}")
        for issue in resp.json():
            if "pull_request" in issue:  # the Issues endpoint also returns pull requests
                skipped_prs += 1
                continue
            issues.append(issue)
        url = resp.links.get("next", {}).get("url")
    return issues, skipped_prs


# --- Microsoft Graph ------------------------------------------------------------------------

class GraphClient:
    def __init__(self):
        require_env("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET")
        self.tenant_id = env("TENANT_ID")
        self.app = msal.ConfidentialClientApplication(
            env("CLIENT_ID"),
            client_credential=env("CLIENT_SECRET"),
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
        )
        self.session = requests.Session()

    def _token(self) -> str:
        # MSAL caches the token and renews it when it expires.
        result = self.app.acquire_token_for_client(scopes=GRAPH_SCOPE)
        if "access_token" not in result:
            raise ConnectorError(
                f"Authentication failed: {result.get('error')}: {result.get('error_description')}"
            )
        return result["access_token"]

    def request(self, method: str, url: str, body=None) -> requests.Response:
        full_url = url if url.startswith("http") else f"{GRAPH}{url}"
        resp = None
        for attempt in range(MAX_RETRIES):
            headers = {"Authorization": f"Bearer {self._token()}", "Content-Type": "application/json"}
            resp = self.session.request(method, full_url, json=body, headers=headers, timeout=120)
            if resp.status_code not in RETRY_STATUSES:
                return resp
            wait = int(resp.headers.get("Retry-After", 2**attempt))
            print(f"  Throttled ({resp.status_code}), retrying in {wait}s...")
            time.sleep(wait)
        return resp


def connection_url() -> str:
    return f"/external/connections/{connection_id()}"


def connection_body(include_id: bool) -> dict:
    body = {
        "name": env("CONNECTION_NAME", "GitHub Issues POC"),
        "description": env(
            "CONNECTION_DESCRIPTION", "GitHub issues (open and closed) from a repository, for Microsoft 365 Copilot."
        ),
    }
    if include_id:
        body = {"id": connection_id(), **body}
    return body


def get_connection_state(graph: GraphClient) -> str | None:
    resp = graph.request("GET", connection_url())
    if resp.status_code == 404:
        return None
    ensure_ok(resp, "Get connection")
    return resp.json().get("state")


def create_connection(graph: GraphClient):
    print(f"Creating connection '{connection_id()}'...")
    resp = graph.request("POST", "/external/connections", connection_body(include_id=True))
    if resp.status_code in (200, 201):
        print("  Connection created.")
    elif resp.status_code == 409:
        print("  Connection already exists; updating name and description.")
        ensure_ok(graph.request("PATCH", connection_url(), connection_body(include_id=False)), "Update connection")
    else:
        ensure_ok(resp, "Create connection")


def register_schema(graph: GraphClient, timeout_s: int = 900):
    if get_connection_state(graph) == "ready":
        print("Schema already registered (connection is ready); skipping.")
        return

    print("Registering schema (this can take several minutes)...")
    resp = graph.request("PATCH", f"{connection_url()}/schema", SCHEMA)
    if resp.status_code not in (200, 202):
        ensure_ok(resp, "Register schema")

    operation_url = resp.headers.get("Location")
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        time.sleep(15)
        if operation_url:
            op = graph.request("GET", operation_url)
            ensure_ok(op, "Check schema operation")
            status = op.json().get("status")
            print(f"  Schema status: {status}")
            if status == "completed":
                return
            if status == "failed":
                raise ConnectorError(f"Schema registration failed:\n{op.text}")
        else:
            state = get_connection_state(graph)
            print(f"  Connection state: {state}")
            if state == "ready":
                return
            if state == "limitExceeded":
                raise ConnectorError("Connection limit exceeded.")
    raise ConnectorError("Timed out waiting for schema registration. Run 'setup' again later.")


def to_external_item(issue: dict, repo: str, tenant_id: str) -> dict:
    author = (issue.get("user") or {}).get("login", "unknown")
    labels = [label["name"] for label in issue.get("labels", [])]
    status = issue["state"]  # "open" or "closed"
    description = (issue.get("body") or "").strip() or "(no description)"

    # The content is full-text indexed, so title, status, URL and labels are also searchable as plain text.
    text = "\n".join(
        [
            f"Issue #{issue['number']}: {issue['title']}",
            f"Repository: {repo}",
            f"Status: {status}",
            f"URL: {issue['html_url']}",
            f"Author: {author}",
            f"Labels: {', '.join(labels) or 'none'}",
            "",
            description,
        ]
    )
    return {
        # Everyone in the tenant can find these issues. Only use this for repositories approved for tenant-wide access.
        "acl": [{"type": "everyone", "value": tenant_id, "accessType": "grant"}],
        "properties": {
            "title": issue["title"],
            "status": status,
            "url": issue["html_url"],
            "issueNumber": issue["number"],
            "author": author,
            "labels": labels,
            "labels@odata.type": "Collection(String)",
            "updatedAt": issue["updated_at"],
        },
        "content": {"type": "text", "value": text},
    }


def ingest(graph: GraphClient, repo: str):
    state = get_connection_state(graph)
    if state != "ready":
        raise ConnectorError(f"Connection state is '{state}', not 'ready'. Run 'setup' first.")

    issues, skipped_prs = fetch_issues(repo)
    print(f"Fetched {len(issues)} issues from {repo} ({skipped_prs} pull requests skipped).")

    ok, failed = 0, 0
    for issue in issues:
        item_id = f"issue{issue['number']}"
        payload = to_external_item(issue, repo, graph.tenant_id)
        resp = graph.request("PUT", f"{connection_url()}/items/{item_id}", payload)
        if resp.ok:
            ok += 1
            print(f"  [OK] {item_id} ({issue['state']}) {issue['title']}")
        else:
            failed += 1
            print(f"  [FAILED] {item_id}: HTTP {resp.status_code} {resp.text}")
    print(f"Done: {ok} indexed, {failed} failed.")
    if failed:
        raise ConnectorError(f"{failed} item(s) failed to ingest")


# --- Commands -------------------------------------------------------------------------------

def cmd_fetch():
    repo = github_repo()
    issues, skipped_prs = fetch_issues(repo)
    open_count = sum(1 for i in issues if i["state"] == "open")
    print(f"{repo}: {len(issues)} issues ({open_count} open, {len(issues) - open_count} closed), "
          f"{skipped_prs} pull requests skipped.")
    for issue in issues[:10]:
        print(f"  #{issue['number']} [{issue['state']}] {issue['title']}")


def cmd_setup():
    connection_id()
    graph = GraphClient()
    create_connection(graph)
    register_schema(graph)
    print("Setup complete.")


def cmd_ingest():
    repo = github_repo()
    connection_id()
    ingest(GraphClient(), repo)


def cmd_all():
    repo = github_repo()
    connection_id()
    graph = GraphClient()
    create_connection(graph)
    register_schema(graph)
    ingest(graph, repo)


def cmd_delete():
    connection_id()
    graph = GraphClient()
    print(f"Deleting connection '{connection_id()}'...")
    resp = graph.request("DELETE", connection_url())
    if resp.status_code in (202, 204):
        print("  Delete requested (indexed items are removed asynchronously).")
    elif resp.status_code == 404:
        print("  Connection not found.")
    else:
        ensure_ok(resp, "Delete connection")


COMMANDS = {"fetch": cmd_fetch, "setup": cmd_setup, "ingest": cmd_ingest, "all": cmd_all, "delete": cmd_delete}


def main() -> int:
    parser = argparse.ArgumentParser(description="Index GitHub Issues into a Microsoft 365 Copilot connector.")
    parser.add_argument("command", choices=COMMANDS)
    args = parser.parse_args()

    load_dotenv()
    try:
        COMMANDS[args.command]()
    except ConnectorError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except requests.RequestException as e:
        print(f"Network error: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
