# GitHub Issues → Microsoft 365 Copilot (Graph synced connector POC)

A single Python script that fetches all open and closed issues from one GitHub repository (pull requests excluded), creates a Microsoft Graph external connection, registers a schema, and ingests the issues so they are searchable in Microsoft 365 Copilot.

## How it works

1. **fetch**: lists issues from the GitHub Issues API (`state=all`, paginated), skipping entries that have a `pull_request` field.
2. **setup**: creates the external connection (`POST /external/connections`) and registers the schema (`PATCH .../schema`), then waits until the connection is `ready`.
3. **ingest**: upserts each issue as an external item (`PUT .../items/issue{number}`). Re-running updates existing items instead of creating duplicates.

Each item contains:

| Field | Stored as | Searchable |
|---|---|---|
| Title | `title` property (labelled as title) | Yes |
| Description (body) | item `content` | Yes (full text) |
| Status (`open` / `closed`) | `status` property and content text | Yes |
| URL | `url` property (labelled as url) | Retrievable |
| Labels, author, issue number, last updated | properties | Labels searchable |

## Prerequisites

- Python 3.10 or later.
- A Microsoft 365 tenant with Copilot connectors available, and a Microsoft 365 Copilot license for the users who will test.
- A Microsoft Entra app registration with a client secret, and the **application** permissions `ExternalConnection.ReadWrite.OwnedBy` and `ExternalItem.ReadWrite.OwnedBy` with **admin consent** granted.
- A GitHub token if the repository is private (fine-grained, read-only **Issues** and **Metadata**).

## Setup

```powershell
cd c:\Fedex_projects\CopilotConnectorGitHubPOC
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
```

Edit `.env`:

- `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`: from the Entra app registration.
- `GITHUB_REPO`: the repository to index, for example `owner/repo`.
- `GITHUB_TOKEN`: required only for private repositories.
- `CONNECTION_ID`: optional. Defaults to `githubissuespoc`. Use a new ID for each connection.

Environment variables can also be set directly in the shell instead of `.env`. Values already set in the shell take precedence over `.env` values.

## Run

Check GitHub access first, without calling Microsoft Graph:

```powershell
python github_issues_connector.py fetch
```

Create the connection, register the schema and ingest all issues in one step:

```powershell
python github_issues_connector.py all
```

Or run the steps separately:

```powershell
python github_issues_connector.py setup    # connection + schema (can take several minutes)
python github_issues_connector.py ingest   # push all issues
```

Remove the connection and its indexed issues when finished:

```powershell
python github_issues_connector.py delete
```

### Create sample GitHub issues from the Excel workbook

To populate `GITHUB_REPO` with the sample issues from `Copilot_GitHub_Issues_Demo_Dataset.xlsx`, first preview them. This makes no GitHub changes:

```powershell
python create_demo_github_issues.py --dry-run
```

Then create them. Issues whose title already exists in the repository are skipped, so you can re-run the command safely. Rows with status `closed` are created and then closed.

```powershell
python create_demo_github_issues.py
```

Use `--file <path>` to read a different workbook. `GITHUB_TOKEN` must have permission to create issues (fine-grained token: **Issues: Read and write**). GitHub cannot set the issue author through the API, so the original author from the workbook is noted at the bottom of each issue body. Labels that don't exist in the repository are expected to be created automatically.

## Verify in Microsoft 365

1. Open the Microsoft 365 admin center and go to **Copilot → Agents and connectors**. Confirm that the connection is listed and enabled.
2. Allow time for indexing. Ingestion succeeding does not mean results appear immediately.
3. In Microsoft 365 Copilot, ask about a known issue, for example:
   - "Find GitHub issues about <a distinctive word from an issue title>."
   - "List closed issues that mention <label name>."
   - "What is the URL of issue #<number> in the connected repository?"

## Limitations and notes

- **Access:** the ACL is `everyone` in the tenant, so every user can find every issue. Use this only for repositories approved for tenant-wide access.
- **Sync is one-way and full:** every `ingest` re-sends all issues. Issues deleted or transferred on GitHub stay indexed until you run `delete`.
- **Comments are not indexed.** Only the title and body are included.
- **Schema changes:** the schema is registered only when the connection is not yet `ready`. To change the schema, run `delete`, then use a new `CONNECTION_ID`.
- **Indexing is asynchronous**, and Copilot decides which sources are relevant to a given question.
- **Limits:** items count against the tenant's external item quota. Very large issue bodies are sent as-is.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `Authentication failed` | Wrong `TENANT_ID`, `CLIENT_ID`, or expired `CLIENT_SECRET`. |
| `HTTP 401` or `403` on connection or item calls | Missing application permissions or admin consent. |
| `HTTP 412` | Tenant is not ready for Copilot connectors (no active Microsoft 365 subscription, or the tenant is too new). |
| `GitHub request failed: HTTP 404` | `GITHUB_REPO` is wrong, or the repository is private and `GITHUB_TOKEN` is missing. |
| `GitHub request failed: HTTP 403` | Rate limit reached, or the token lacks access. Add or fix `GITHUB_TOKEN`. |
| `Connection state is 'draft'` during `ingest` | Run `setup` first. |

## Files

- [github_issues_connector.py](github_issues_connector.py): the complete application.
- [create_demo_github_issues.py](create_demo_github_issues.py): creates GitHub issues from the demo Excel workbook. It reuses config helpers from the connector.
- [Copilot_GitHub_Issues_Demo_Dataset.xlsx](Copilot_GitHub_Issues_Demo_Dataset.xlsx): fictional sample issues and demo scenarios.
- [requirements.txt](requirements.txt): Python dependencies.
- [.env.example](.env.example): configuration template.
