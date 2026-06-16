# Dynatrace Roles

A simple tool for creating and assigning Dynatrace user roles.

## Getting started

The tool automates the following actions:
- creation of management zones,
- creation of user groups,
- updating the permissions of user groups,
- adding users to groups.

Inputs for the tool are two sheets in one Excel file:
- role definitions table,
- users list table.

The tool supports user roles with management zone (MZ), environment and cluster level permissions.
For MZ level roles it creates one user group per management zone, for env/cluster levels - one user group per Dynatrace cluster.

Users should exist in Dynatrace cluster with the specified user IDs before running the tool.

Management zones are created with 3 rules: include hosts, processes and services filtered by the specified tag.

### Installing

This tool is written in Python, so you need [python3](https://www.python.org/downloads/) and [pip](https://pip.pypa.io/en/stable/installing/) installed.
To install the tool prerequisites simply download the repository and run:

```bash
pip install -r requirements.txt
```

### Configuring

Before running the tool, you need to prepare the input Excel file (e.g., [users.xlsx]users.xlsx)) containing two sheets: one for role definitions and another for user assignments.

#### 1. Role Definitions Sheet (e.g., `roles_test` / `roles_prod`)

This sheet defines the permissions associated with each role name.

- **First row**: Specifies permission levels (`cluster`, `env`, or `mz`). The first cell (A1) should be empty.
- **Second row**: Specifies specific Dynatrace permission codes (e.g., `VIEWER`, `MANAGE_SETTINGS`, `AGENT_INSTALL`, etc.). The first cell (A2) should be empty.
- **Subsequent rows**:
  - Column A: The unique name of the role (e.g., `IT Manager`, `Business User Process Owner`).
  - Remaining columns: Mark with `+` (or any non-empty value) to grant the permission, or leave blank to deny.

Example structure:

| | cluster | env | env | env | mz | mz |
|---|---|---|---|---|---|---|
| | | VIEWER | MANAGE_SETTINGS | LOG_VIEWER | VIEWER | MANAGE_SETTINGS |
| **IT Manager** | | + | | | | |
| **AS Administrator** | | + | + | + | | |
| **Business User** | | | | | + | |
| **Global Admin** | + | | | | | |

*Note: For `env` and `mz` roles, the `VIEWER` permission must be enabled if any other permission is enabled.*

#### 2. User Assignments Sheet (e.g., `users`)

This sheet maps users to roles and specifies optional management zone configuration.

- **Header row**: Must contain at least 5 columns:
  - **User ID**: The Dynatrace user identifier (the user must already exist in Dynatrace).
  - **Role Name**: The role name, matching one of the roles defined in the Roles sheet.
  - **Tag key**: (Optional) Dynatrace tag key to filter hosts, processes, and services by.
  - **Tag value**: (Optional) Dynatrace tag value to filter by.
  - **MZ name parts**: (Optional, columns 5 and beyond) One or more columns containing parts of the management zone name. They are concatenated with `' - '` to form the full management zone name.

Example structure:

| User ID | Role Name | Tag key | Tag value | MZ name part 1 | MZ name part 2 |
|---|---|---|---|---|---|
| user1 | IT Manager | | | | |
| user2 | Business User | IT Service CI | CI0000001 | IT Service 3 | CI0000001 |
| user3 | Business User | IT Service CI | CI0000002 | IT Service 3 | CI0000002 |

*Note: For MZ-level roles, the tool will automatically create a management zone (if it doesn't exist) named by joining the MZ name parts (e.g., `IT Service 3 - CI0000001`) with rules matching the specified tags. It will also create a corresponding user group named `{Role Name} - {MZ Name}` (e.g., `Business User - IT Service 3 - CI0000001`).*

---

### Running the tool

By default, the tool runs in **dry-run / test mode**, showing what changes it would make without actually applying them. Use the `--apply-changes` flag to commit changes to Dynatrace.

```bash
python dt-roles.py --cluster-hostname <hostname> \
                   --cluster-token <token> \
                   --env-id <env_id> \
                   --env-token <env_token> \
                   --file users.xlsx \
                   --users-sheet users \
                   --roles-sheet roles_test
```

#### Command-line arguments:

- `--cluster-hostname`: Dynatrace cluster hostname (required).
- `--cluster-token`: Dynatrace API cluster token (required).
- `--env-id`: Dynatrace environment ID (required).
- `--env-token`: Dynatrace API environment token (required).
- `--file`: Path to the input Excel file (required).
- `--users-sheet`: Sheet name for the users list (required).
- `--roles-sheet`: Sheet name for the role definitions (required).
- `--apply-changes`: Apply changes to Dynatrace configuration (dry-run by default).

