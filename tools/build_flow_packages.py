#!/usr/bin/env python3
"""
build_flow_packages.py — generate importable LEGACY Power Automate packages (.zip)
for every lab flow, plus the mock-data workbooks each lab needs.

A legacy package is the format produced by "Export > Package (.zip)" in Power Automate
and consumed by "Import > Import Package (Legacy)". It contains:

    manifest.json          package metadata + the resource graph
    <FlowName>/
        definition.json    the flow definition (trigger + actions)
        apisMap.json       connector references

The definitions here are the SAME flows built live in the course environment, so a
trainer can either use the pre-built flows or import these packages into a fresh
environment and get an identical starting point.
"""
import json, os, shutil, sys, zipfile, uuid, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
LABS = os.path.join(REPO, "labs")

NOW = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S.0000000Z")

O365 = "shared_office365"
EXCEL = "shared_excelonlinebusiness"
APPROVALS = "shared_approvals"
POWERAPPS = "shared_powerappsforappmakers"


def conn_ref(name, api):
    return {
        "connectionName": name,
        "source": "Embedded",
        "id": f"/providers/Microsoft.PowerApps/apis/{api}",
        "tier": "NotSpecified",
    }


def wrap(display_name, definition, connections):
    """Wrap a flow definition into the legacy package structure."""
    return {
        "properties": {
            "connectionReferences": connections,
            "definition": definition,
            "displayName": display_name,
        }
    }


def base_def(triggers, actions, params=None):
    d = {
        "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
        "contentVersion": "1.0.0.0",
        "parameters": params or {},
        "triggers": triggers,
        "actions": actions,
        "outputs": {},
    }
    return d


def o365_conn_param():
    return {
        "$connections": {"defaultValue": {}, "type": "Object"},
        "$authentication": {"defaultValue": {}, "type": "SecureObject"},
    }


def api_host(api, operation):
    return {
        "host": {
            "connectionName": api.replace("shared_", "shared_"),
            "operationId": operation,
            "apiId": f"/providers/Microsoft.PowerApps/apis/{api}",
        }
    }


def openapi(api, operation, parameters, run_after=None, conn_key=None):
    a = {
        "type": "OpenApiConnection",
        "inputs": {
            "parameters": parameters,
            "host": {
                "apiId": f"/providers/Microsoft.PowerApps/apis/{api}",
                "connection": conn_key or api,
                "operationId": operation,
            },
        },
        "runAfter": run_after or {},
    }
    return a


# --------------------------------------------------------------------------- flows
def flow_lab3():
    """Lab 3 — Trigger and Actions: email trigger -> Compose -> Send an email."""
    triggers = {
        "When_a_new_email_arrives_(V3)": {
            "type": "OpenApiConnectionNotification",
            "inputs": {
                "parameters": {"folderPath": "Inbox", "subjectFilter": "SERVICE",
                               "importance": "Any", "includeAttachments": False},
                "host": {"apiId": f"/providers/Microsoft.PowerApps/apis/{O365}",
                         "connection": O365, "operationId": "OnNewEmailV3"},
            },
            "splitOn": "@triggerOutputs()?['body/value']",
        }
    }
    actions = {
        "Compose": {
            "type": "Compose",
            "inputs": ("Service call logged | From: @{triggerOutputs()?['body/from']}"
                       " | Subject: @{triggerOutputs()?['body/subject']}"
                       " | Received: @{triggerOutputs()?['body/receivedDateTime']}"),
            "runAfter": {},
        },
        "Send_an_email_(V2)": openapi(
            O365, "SendEmailV2",
            {"emailMessage/To": "admin@tertiaryinfotech.onmicrosoft.com",
             "emailMessage/Subject": "Service call logged",
             "emailMessage/Body": "<p>A new service call was received:</p><p>@{outputs('Compose')}</p>",
             "emailMessage/Importance": "Normal"},
            run_after={"Compose": ["Succeeded"]}),
    }
    return base_def(triggers, actions, o365_conn_param())


def flow_lab4():
    """Lab 4 — Log to Excel: email trigger -> Compose -> add row -> notify."""
    d = flow_lab3()
    d["actions"]["Add_a_row_into_a_table"] = openapi(
        EXCEL, "AddRowV2",
        {"source": "me", "drive": "OneDrive", "file": "/KinetEco Service Calls.xlsx",
         "table": "ServiceCalls",
         "item/Reported By": "@triggerOutputs()?['body/from']",
         "item/Problem": "@triggerOutputs()?['body/subject']",
         "item/Date Reported": "@triggerOutputs()?['body/receivedDateTime']"},
        run_after={"Compose": ["Succeeded"]})
    d["actions"]["Send_an_email_(V2)"]["runAfter"] = {"Add_a_row_into_a_table": ["Succeeded"]}
    return d


def flow_lab5():
    """Lab 5 — Conditions and Branching: urgent escalates, routine only logs."""
    d = flow_lab4()
    log = d["actions"].pop("Add_a_row_into_a_table")
    mail = d["actions"].pop("Send_an_email_(V2)")
    log["runAfter"] = {}
    mail["runAfter"] = {}
    mail["inputs"]["parameters"]["emailMessage/Importance"] = "High"
    mail["inputs"]["parameters"]["emailMessage/Subject"] = "URGENT service call"
    d["actions"]["Condition"] = {
        "type": "If",
        "expression": {
            "contains": ["@toupper(triggerOutputs()?['body/subject'])", "URGENT"]
        },
        "actions": {"Escalate_to_duty_engineer": mail, "Log_urgent_call": json.loads(json.dumps(log))},
        "else": {"actions": {"Log_routine_call": json.loads(json.dumps(log))}},
        "runAfter": {"Compose": ["Succeeded"]},
    }
    return d


def flow_lab6():
    """Lab 6 — Leave Application Approval: manual trigger -> approval -> branch."""
    triggers = {
        "manual": {
            "type": "Request",
            "kind": "Button",
            "inputs": {"schema": {"type": "object", "properties": {
                "Applicant": {"type": "string", "title": "Applicant"},
                "LeaveType": {"type": "string", "title": "Leave Type"},
                "StartDate": {"type": "string", "title": "Start Date"},
                "Days": {"type": "number", "title": "Days"},
                "Reason": {"type": "string", "title": "Reason"}},
                "required": ["Applicant", "LeaveType", "StartDate", "Days"]}},
        }
    }
    approve_mail = openapi(
        O365, "SendEmailV2",
        {"emailMessage/To": "@triggerBody()?['Applicant']",
         "emailMessage/Subject": "Leave request APPROVED",
         "emailMessage/Body": "<p>Your leave request has been approved.</p>"},
        run_after={})
    reject_mail = openapi(
        O365, "SendEmailV2",
        {"emailMessage/To": "@triggerBody()?['Applicant']",
         "emailMessage/Subject": "Leave request REJECTED",
         "emailMessage/Body": "<p>Your request was rejected. Comments: @{body('Start_and_wait_for_an_approval')?['responses'][0]['comments']}</p>"},
        run_after={})
    actions = {
        "Start_and_wait_for_an_approval": openapi(
            APPROVALS, "StartAndWaitForAnApproval",
            {"approvalType": "Basic",
             "ApprovalCreationInput/title": "Leave request from @{triggerBody()?['Applicant']}",
             "ApprovalCreationInput/assignedTo": "admin@tertiaryinfotech.onmicrosoft.com",
             "ApprovalCreationInput/details":
                 "@{triggerBody()?['Days']} day(s) of @{triggerBody()?['LeaveType']} from @{triggerBody()?['StartDate']}. Reason: @{triggerBody()?['Reason']}"},
            run_after={}),
        "Condition": {
            "type": "If",
            "expression": {"equals": ["@body('Start_and_wait_for_an_approval')?['outcome']", "Approve"]},
            "actions": {"Notify_approved": approve_mail},
            "else": {"actions": {"Notify_rejected": reject_mail}},
            "runAfter": {"Start_and_wait_for_an_approval": ["Succeeded"]},
        },
    }
    return base_def(triggers, actions, o365_conn_param())


def flow_lab11():
    """Lab 11 — Submit Leave Request, called from the canvas app (PowerApps V2)."""
    triggers = {
        "PowerApps_(V2)": {
            "type": "Request",
            "kind": "PowerApp",
            "inputs": {"schema": {"type": "object", "properties": {
                "Applicant": {"type": "string"}, "LeaveType": {"type": "string"},
                "StartDate": {"type": "string"}, "Days": {"type": "number"},
                "Reason": {"type": "string"}}}},
        }
    }
    actions = {
        "Send_an_email_(V2)": openapi(
            O365, "SendEmailV2",
            {"emailMessage/To": "admin@tertiaryinfotech.onmicrosoft.com",
             "emailMessage/Subject": "New leave request from @{triggerBody()['Applicant']}",
             "emailMessage/Body": "<p>@{triggerBody()['Days']} day(s) of @{triggerBody()['LeaveType']} from @{triggerBody()['StartDate']}</p><p>Reason: @{triggerBody()['Reason']}</p>"},
            run_after={}),
        "Add_a_row_into_a_table": openapi(
            EXCEL, "AddRowV2",
            {"source": "me", "drive": "OneDrive", "file": "/LeaveLog.xlsx", "table": "LeaveLog",
             "item/Applicant": "@triggerBody()['Applicant']",
             "item/LeaveType": "@triggerBody()['LeaveType']",
             "item/StartDate": "@triggerBody()['StartDate']",
             "item/Days": "@triggerBody()['Days']",
             "item/Status": "Submitted"},
            run_after={"Send_an_email_(V2)": ["Succeeded"]}),
    }
    return base_def(triggers, actions, o365_conn_param())


def flow_lab12():
    """Lab 12 — Return Leave Balance: closes the round trip back to the app."""
    triggers = {
        "PowerApps_(V2)": {
            "type": "Request", "kind": "PowerApp",
            "inputs": {"schema": {"type": "object", "properties": {
                "Applicant": {"type": "string"}, "Days": {"type": "number"}}}},
        }
    }
    actions = {
        "Compose_balance": {
            "type": "Compose",
            "inputs": "@sub(14, int(triggerBody()['Days']))",
            "runAfter": {},
        },
        "Compose_reference": {
            "type": "Compose",
            "inputs": "@concat('LV-', formatDateTime(utcNow(),'yyyyMMdd'), '-', rand(1000,9999))",
            "runAfter": {"Compose_balance": ["Succeeded"]},
        },
        "Respond_to_a_Power_App_or_flow": {
            "type": "Response",
            "kind": "PowerApp",
            "inputs": {
                "statusCode": 200,
                "body": {
                    "balance": "@outputs('Compose_balance')",
                    "reference": "@outputs('Compose_reference')",
                    "status": "Submitted",
                },
                "schema": {"type": "object", "properties": {
                    "balance": {"type": "number"}, "reference": {"type": "string"},
                    "status": {"type": "string"}}},
            },
            "runAfter": {"Compose_reference": ["Succeeded"]},
        },
    }
    return base_def(triggers, actions, o365_conn_param())


def flow_lab13():
    """Lab 13 — Approval Round Trip: app -> approval -> write back -> respond."""
    triggers = {
        "PowerApps_(V2)": {
            "type": "Request", "kind": "PowerApp",
            "inputs": {"schema": {"type": "object", "properties": {
                "Applicant": {"type": "string"}, "LeaveType": {"type": "string"},
                "StartDate": {"type": "string"}, "Days": {"type": "number"},
                "Reason": {"type": "string"}}}},
        }
    }
    actions = {
        "Add_pending_row": openapi(
            EXCEL, "AddRowV2",
            {"source": "me", "drive": "OneDrive", "file": "/LeaveLog.xlsx", "table": "LeaveLog",
             "item/Applicant": "@triggerBody()['Applicant']",
             "item/LeaveType": "@triggerBody()['LeaveType']",
             "item/StartDate": "@triggerBody()['StartDate']",
             "item/Days": "@triggerBody()['Days']",
             "item/Status": "Pending"},
            run_after={}),
        "Start_and_wait_for_an_approval": openapi(
            APPROVALS, "StartAndWaitForAnApproval",
            {"approvalType": "Basic",
             "ApprovalCreationInput/title": "Leave request from @{triggerBody()['Applicant']}",
             "ApprovalCreationInput/assignedTo": "admin@tertiaryinfotech.onmicrosoft.com",
             "ApprovalCreationInput/details":
                 "@{triggerBody()['Days']} day(s) of @{triggerBody()['LeaveType']} from @{triggerBody()['StartDate']}"},
            run_after={"Add_pending_row": ["Succeeded"]}),
        "Condition": {
            "type": "If",
            "expression": {"equals": ["@body('Start_and_wait_for_an_approval')?['outcome']", "Approve"]},
            "actions": {"Set_approved": {
                "type": "Compose", "inputs": "Approved", "runAfter": {}}},
            "else": {"actions": {"Set_rejected": {
                "type": "Compose", "inputs": "Rejected", "runAfter": {}}}},
            "runAfter": {"Start_and_wait_for_an_approval": ["Succeeded"]},
        },
        "Respond_to_a_Power_App_or_flow": {
            "type": "Response", "kind": "PowerApp",
            "inputs": {
                "statusCode": 200,
                "body": {
                    "status": "@body('Start_and_wait_for_an_approval')?['outcome']",
                    "comments": "@{body('Start_and_wait_for_an_approval')?['responses'][0]['comments']}",
                },
                "schema": {"type": "object", "properties": {
                    "status": {"type": "string"}, "comments": {"type": "string"}}},
            },
            "runAfter": {"Condition": ["Succeeded"]},
        },
    }
    return base_def(triggers, actions, o365_conn_param())


def flow_lab5b():
    """Lab 5b — Daily Digest: scheduled recurrence -> list rows -> HTML table -> email."""
    triggers = {
        "Recurrence": {
            "type": "Recurrence",
            "recurrence": {"frequency": "Day", "interval": 1,
                           "timeZone": "Singapore Standard Time",
                           "schedule": {"hours": ["18"], "minutes": [0]}},
        }
    }
    actions = {
        "List_rows_present_in_a_table": openapi(
            EXCEL, "GetItems",
            {"source": "me", "drive": "OneDrive",
             "file": "/KinetEco Service Calls.xlsx", "table": "ServiceCalls"},
            run_after={}),
        "Create_HTML_table": {
            "type": "Table",
            "inputs": {"from": "@outputs('List_rows_present_in_a_table')?['body/value']",
                       "format": "HTML"},
            "runAfter": {"List_rows_present_in_a_table": ["Succeeded"]},
        },
        "Send_an_email_(V2)": openapi(
            O365, "SendEmailV2",
            {"emailMessage/To": "admin@tertiaryinfotech.onmicrosoft.com",
             "emailMessage/Subject": "Daily service call digest",
             "emailMessage/Body": "<p>Service calls logged today:</p>@{body('Create_HTML_table')}"},
            run_after={"Create_HTML_table": ["Succeeded"]}),
    }
    return base_def(triggers, actions, o365_conn_param())


FLOWS = [
    # (lab number, flow display name, builder, connectors used)
    (3,  "Lab 3 - Trigger and Actions (DO NOT DELETE)",        flow_lab3,  [O365]),
    (4,  "Lab 4 - Log to Excel (DO NOT DELETE)",               flow_lab4,  [O365, EXCEL]),
    (5,  "Lab 5 - Conditions and Branching (DO NOT DELETE)",   flow_lab5,  [O365, EXCEL]),
    (5,  "Lab 5b - Daily Digest (DO NOT DELETE)",              flow_lab5b, [O365, EXCEL]),
    (6,  "Lab 6 - Leave Application Approval (DO NOT DELETE)", flow_lab6,  [O365, APPROVALS]),
    (11, "Lab 11 - Submit Leave Request (DO NOT DELETE)",      flow_lab11, [O365, EXCEL]),
    (12, "Lab 12 - Return Leave Balance (DO NOT DELETE)",      flow_lab12, [O365]),
    (13, "Lab 13 - Approval Round Trip (DO NOT DELETE)",       flow_lab13, [O365, EXCEL, APPROVALS]),
]

API_DISPLAY = {
    O365: "Office 365 Outlook",
    EXCEL: "Excel Online (Business)",
    APPROVALS: "Approvals",
    POWERAPPS: "Power Apps for Makers",
}


def build_package(display_name, definition, apis, out_zip):
    """Assemble a legacy Power Automate package (.zip)."""
    flow_guid = str(uuid.uuid4())
    res_key = str(uuid.uuid4()).replace("-", "")[:8].upper()

    connection_refs = {}
    for api in apis:
        connection_refs[api] = {
            "connectionName": api,
            "source": "Embedded",
            "id": f"/providers/Microsoft.PowerApps/apis/{api}",
            "tier": "NotSpecified",
        }

    # ---- manifest.json
    resources = {
        res_key: {
            "id": flow_guid,
            "name": flow_guid,
            "type": "Microsoft.Flow/flows",
            "suggestedCreationType": "New",
            "creationType": "New, Existing, Update",
            "details": {"displayName": display_name},
            "configurableBy": "User",
            "hierarchy": "Root",
            "dependsOn": [],
        }
    }
    for i, api in enumerate(apis):
        k = f"{res_key}{i+1}"
        resources[res_key]["dependsOn"].append(k)
        resources[k] = {
            "id": f"/providers/Microsoft.PowerApps/apis/{api}",
            "name": api,
            "type": "Microsoft.PowerApps/apis",
            "suggestedCreationType": "Existing",
            "creationType": "Existing",
            "details": {"displayName": API_DISPLAY.get(api, api)},
            "configurableBy": "System",
            "hierarchy": "Child",
            "dependsOn": [],
        }

    manifest = {
        "schema": "1.0",
        "details": {
            "displayName": display_name,
            "description": f"{display_name} — Tertiary Infotech Academy, TGS-2022015539",
            "createdTime": NOW,
            "packageTelemetryId": str(uuid.uuid4()),
            "creator": "Tertiary Infotech Academy",
            "sourceEnvironment": "TGS-2022015539-Applications Integration with Power Apps and Power Automate",
        },
        "resources": resources,
    }

    # ---- the flow resource folder
    safe = display_name.replace("/", "-")
    flow_json = {
        "name": flow_guid,
        "id": f"/providers/Microsoft.Flow/flows/{flow_guid}",
        "type": "Microsoft.Flow/flows",
        "properties": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
            "displayName": display_name,
            "definition": definition,
            "connectionReferences": connection_refs,
        },
    }

    os.makedirs(os.path.dirname(out_zip), exist_ok=True)
    if os.path.exists(out_zip):
        os.remove(out_zip)
    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("manifest.json", json.dumps(manifest, indent=2))
        z.writestr(f"Microsoft.Flow/flows/{flow_guid}/definition.json",
                   json.dumps(flow_json, indent=2))
        z.writestr(f"Microsoft.Flow/flows/{flow_guid}/apisMap.json",
                   json.dumps({api: f"/providers/Microsoft.PowerApps/apis/{api}"
                               for api in apis}, indent=2))
    return out_zip


def main():
    made = []
    for lab, name, builder, apis in FLOWS:
        folder = os.path.join(LABS, f"lab-{lab:02d}")
        os.makedirs(folder, exist_ok=True)
        slug = name.split(" (DO NOT")[0].replace(" ", "-").replace("---", "-")
        out = os.path.join(folder, f"{slug}.zip")
        build_package(name, builder(), apis, out)
        made.append(out)
        # also drop the raw definition so learners can read it
        with open(os.path.join(folder, f"{slug}-definition.json"), "w") as f:
            json.dump(builder(), f, indent=2)
    print(f"built {len(made)} package(s):")
    for m in made:
        print("  ", os.path.relpath(m, REPO), os.path.getsize(m), "bytes")


if __name__ == "__main__":
    main()
