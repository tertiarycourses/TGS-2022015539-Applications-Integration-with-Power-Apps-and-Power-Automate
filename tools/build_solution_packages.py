#!/usr/bin/env python3
"""
build_solution_packages.py — generate importable DATAVERSE SOLUTION packages (.zip)
for the lab flows.

The structure here is copied EXACTLY from a real solution exported from the course
environment (tools/reference/), because the solution importer is strict:
  * RootComponent ids are LOWERCASE and must equal the WorkflowId
  * the Workflow element needs RunAs / IsTransacted / ModernFlowType / etc.
  * [Content_Types].xml declares xml as application/octet-stream
  * the workflow JSON uses connectionReferences with api/connection/runtimeSource

The tenant has "Create in Dataverse solutions" enabled, which disables legacy package
import, so this is the format that actually imports here. Every lab folder also ships
the legacy .zip for portability to tenants without that policy.
"""
import json, os, uuid, zipfile, html, re

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
LABS = os.path.join(REPO, "labs")
REF = os.path.join(HERE, "reference")

import importlib.util
spec = importlib.util.spec_from_file_location("bfp", os.path.join(HERE, "build_flow_packages.py"))
bfp = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bfp)

CONTENT_TYPES = ('﻿<?xml version="1.0" encoding="utf-8"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="xml" ContentType="application/octet-stream" />'
                 '<Default Extension="json" ContentType="application/octet-stream" />'
                 '</Types>')

# The publisher that exists in this environment (from the real export).
PUBLISHER_UNIQUE = "Cr8e22a"
PUBLISHER_DISPLAY = "CDS Default Publisher"
PUBLISHER_PREFIX = "cr8e22a"

# Only the Office 365 Outlook connection reference actually exists in the course
# environment (created when Lab 3 was built). Any other connector must bind at
# invoke time instead of naming a reference that does not exist.
CONNREF_LOGICAL = {
    bfp.O365: "new_sharedoffice365_2bc04",
}


def solution_xml(unique_name, display_name, version, root_ids):
    roots = "\n".join(
        f'      <RootComponent type="29" id="{{{g.lower()}}}" behavior="0" />' for g in root_ids)
    return f"""<ImportExportXml version="9.2.26082.148" SolutionPackageVersion="9.2" languagecode="1033" generatedBy="CrmLive" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" OrganizationVersion="9.2.26082.148" OrganizationSchemaType="Standard" CRMServerServiceabilityVersion="9.2.26083.00157">
  <SolutionManifest>
    <UniqueName>{unique_name}</UniqueName>
    <LocalizedNames>
      <LocalizedName description="{html.escape(display_name)}" languagecode="1033" />
    </LocalizedNames>
    <Descriptions />
    <Version>{version}</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>{PUBLISHER_UNIQUE}</UniqueName>
      <LocalizedNames>
        <LocalizedName description="{PUBLISHER_DISPLAY}" languagecode="1033" />
      </LocalizedNames>
      <Descriptions />
      <EMailAddress xsi:nil="true"></EMailAddress>
      <SupportingWebsiteUrl xsi:nil="true"></SupportingWebsiteUrl>
      <CustomizationPrefix>{PUBLISHER_PREFIX}</CustomizationPrefix>
      <CustomizationOptionValuePrefix>10000</CustomizationOptionValuePrefix>
      <Addresses></Addresses>
    </Publisher>
    <RootComponents>
{roots}
    </RootComponents>
    <MissingDependencies />
  </SolutionManifest>
</ImportExportXml>"""


def workflow_element(name, guid, json_name):
    return f"""    <Workflow WorkflowId="{{{guid.lower()}}}" Name="{html.escape(name)}">
      <JsonFileName>/Workflows/{json_name}</JsonFileName>
      <Type>1</Type>
      <Subprocess>0</Subprocess>
      <Category>5</Category>
      <Mode>0</Mode>
      <Scope>4</Scope>
      <OnDemand>0</OnDemand>
      <TriggerOnCreate>0</TriggerOnCreate>
      <TriggerOnDelete>0</TriggerOnDelete>
      <AsyncAutodelete>0</AsyncAutodelete>
      <SyncWorkflowLogOnFailure>0</SyncWorkflowLogOnFailure>
      <StateCode>1</StateCode>
      <StatusCode>2</StatusCode>
      <RunAs>1</RunAs>
      <IsTransacted>1</IsTransacted>
      <IntroducedVersion>1.0</IntroducedVersion>
      <IsCustomizable>1</IsCustomizable>
      <BusinessProcessType>0</BusinessProcessType>
      <IsCustomProcessingStepAllowedForOtherPublishers>1</IsCustomProcessingStepAllowedForOtherPublishers>
      <ModernFlowType>0</ModernFlowType>
      <PrimaryEntity>none</PrimaryEntity>
      <LocalizedNames>
        <LocalizedName languagecode="1033" description="{html.escape(name)}" />
      </LocalizedNames>
    </Workflow>"""


def customizations_xml(workflows):
    body = "\n".join(workflow_element(n, g, f) for n, g, f in workflows)
    return f"""<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <Entities></Entities>
  <Roles></Roles>
  <Workflows>
{body}
  </Workflows>
  <FieldSecurityProfiles></FieldSecurityProfiles>
  <Templates></Templates>
  <EntityMaps></EntityMaps>
  <EntityRelationships></EntityRelationships>
  <OrganizationSettings></OrganizationSettings>
  <optionsets></optionsets>
  <CustomControls></CustomControls>
  <EntityDataProviders></EntityDataProviders>
  <Languages>
    <Language>1033</Language>
  </Languages>
</ImportExportXml>"""


def workflow_json(definition, apis):
    # "invoker" binds the connection at import/first-open time, so the package does not
    # depend on connection references that happen to exist in the source environment.
    # Only shared_office365 has a real connection reference here; everything else would
    # fail import with "Failed to find connection references with logical name(s) ...".
    refs = {}
    for api in apis:
        logical = CONNREF_LOGICAL.get(api)
        if logical:
            refs[api] = {
                "api": {"name": api},
                "connection": {"connectionReferenceLogicalName": logical},
                "runtimeSource": "embedded",
            }
        else:
            refs[api] = {"api": {"name": api}, "runtimeSource": "invoker"}
    # host blocks in a solution flow carry connectionName, matching the real export
    d = json.loads(json.dumps(definition))

    def fix_hosts(node):
        if isinstance(node, dict):
            if "host" in node and isinstance(node["host"], dict) and "apiId" in node["host"]:
                h = node["host"]
                api = h["apiId"].rsplit("/", 1)[-1]
                node["host"] = {"apiId": h["apiId"],
                                "operationId": h.get("operationId"),
                                "connectionName": api}
                # in a solution flow the connector inputs put host FIRST
                if "parameters" in node:
                    params = node.pop("parameters")
                    host = node.pop("host")
                    node["host"] = host
                    node["parameters"] = params
            for v in node.values():
                fix_hosts(v)
        elif isinstance(node, list):
            for v in node:
                fix_hosts(v)

    fix_hosts(d)
    # Solution flows carry schemaVersion at the TOP level and templateName in properties.
    return {"properties": {"connectionReferences": refs, "definition": d,
                           "templateName": None},
            "schemaVersion": "1.0.0.0"}


def safe_name(display):
    return re.sub(r"[^A-Za-z0-9]", "", display)


def build_solution(unique_name, display_name, flows, out_zip, version="1.0.0.0"):
    """flows: list of (display_name, definition, apis)"""
    workflows, files, root_ids = [], {}, []
    for disp, definition, apis in flows:
        guid = str(uuid.uuid4()).upper()
        json_name = f"{safe_name(disp)}-{guid}.json"
        files[f"Workflows/{json_name}"] = json.dumps(workflow_json(definition, apis), indent=2)
        workflows.append((disp, guid, json_name))
        root_ids.append(guid)

    os.makedirs(os.path.dirname(out_zip), exist_ok=True)
    if os.path.exists(out_zip):
        os.remove(out_zip)
    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("solution.xml", solution_xml(unique_name, display_name, version, root_ids))
        z.writestr("customizations.xml", customizations_xml(workflows))
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        for k, v in files.items():
            z.writestr(k, v)
    return out_zip


def main():
    by_lab = {}
    for lab, name, builder, apis in bfp.FLOWS:
        by_lab.setdefault(lab, []).append((name, builder(), apis))

    made = []
    for lab, flows in sorted(by_lab.items()):
        folder = os.path.join(LABS, f"lab-{lab:02d}")
        os.makedirs(folder, exist_ok=True)
        out = os.path.join(folder, f"Solution-Lab-{lab:02d}.zip")
        build_solution(f"TGS2022015539Lab{lab:02d}",
                       f"TGS-2022015539 Lab {lab} (DO NOT DELETE)", flows, out)
        made.append(out)

    all_flows = [(n, b(), a) for _, n, b, a in bfp.FLOWS]
    combined = os.path.join(LABS, "Solution-All-Labs.zip")
    build_solution("TGS2022015539AllLabsPkg",
                   "TGS-2022015539 All Labs Package (DO NOT DELETE)", all_flows, combined)
    made.append(combined)

    print(f"built {len(made)} solution package(s):")
    for m in made:
        print("  ", os.path.relpath(m, REPO), os.path.getsize(m), "bytes")


if __name__ == "__main__":
    main()
