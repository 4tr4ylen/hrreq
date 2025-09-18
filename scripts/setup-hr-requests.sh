#!/usr/bin/env bash
set -euo pipefail

# Prereq: CLI for Microsoft 365
#   npm i -g @pnp/cli-microsoft365 --prefix "$HOME/.npm-global"
#   export PATH="$HOME/.npm-global/bin:$PATH"
#   m365 login --authType deviceCode

SITE_URL="https://atranox.sharepoint.com/sites/test"
LIST_TITLE="HR Requests"
DEPARTMENTS=("HR" "IT" "Finance" "Sales" "Operations")
REQUEST_TYPES=("Leave Request" "Equipment Request" "Policy Question" "Benefits Question" "Other")

usage() {
  echo "Usage: $0 [-s siteUrl] [-l listTitle]" >&2
  exit 1
}

while getopts "s:l:" opt; do
  case $opt in
    s) SITE_URL="$OPTARG" ;;
    l) LIST_TITLE="$OPTARG" ;;
    *) usage ;;
  esac
done

echo "Using site: $SITE_URL"
echo "List title: $LIST_TITLE"

# Create list if not exists
if ! m365 spo list get --webUrl "$SITE_URL" --title "$LIST_TITLE" >/dev/null 2>&1; then
  echo "Creating list '$LIST_TITLE'"
  m365 spo list add --webUrl "$SITE_URL" --title "$LIST_TITLE" --baseTemplate GenericList >/dev/null
fi

echo "Enable attachments and versioning"
m365 spo list set --webUrl "$SITE_URL" --title "$LIST_TITLE" --enableAttachments true --enableVersioning true --majorVersionLimit 500 >/dev/null

# Add/ensure fields via XML (idempotent)
add_field_xml() {
  local xml="$1"
  # Try add; ignore if exists
  m365 spo field add --webUrl "$SITE_URL" --listTitle "$LIST_TITLE" --xml "$xml" >/dev/null 2>&1 || true
}

echo "Ensuring fields"

# Description (Note)
add_field_xml '<Field DisplayName="Description" Name="Description" Type="Note" Required="TRUE" NumLines="6" />'

# Department (Choice)
dept_choices=$(printf '<CHOICES>%s</CHOICES>' "$(printf '<CHOICE>%s</CHOICE>' "${DEPARTMENTS[@]}")")
add_field_xml "<Field DisplayName=\"Department\" Name=\"Department\" Type=\"Choice\" Required=\"TRUE\">${dept_choices}</Field>"

# Request Type (Choice)
rt_choices=$(printf '<CHOICES>%s</CHOICES>' "$(printf '<CHOICE>%s</CHOICE>' "${REQUEST_TYPES[@]}")")
add_field_xml "<Field DisplayName=\"Request Type\" Name=\"RequestType\" Type=\"Choice\" Required=\"TRUE\">${rt_choices}<Default>Other</Default></Field>"

# Person fields
add_field_xml '<Field DisplayName="Requestor" Name="Requestor" Type="User" UserSelectionMode="PeopleOnly" Required="FALSE" />'
add_field_xml '<Field DisplayName="Manager" Name="Manager" Type="User" UserSelectionMode="PeopleOnly" Required="FALSE" />'

# Status
add_field_xml '<Field DisplayName="Status" Name="Status" Type="Choice" Required="TRUE"><Default>Submitted</Default><CHOICES><CHOICE>Draft</CHOICE><CHOICE>Submitted</CHOICE><CHOICE>Pending Approval</CHOICE><CHOICE>Approved</CHOICE><CHOICE>Rejected</CHOICE><CHOICE>Completed</CHOICE></CHOICES></Field>'

# ApprovalOutcome
add_field_xml '<Field DisplayName="Approval Outcome" Name="ApprovalOutcome" Type="Choice" Required="FALSE"><CHOICES><CHOICE>Approved</CHOICE><CHOICE>Rejected</CHOICE></CHOICES></Field>'

# Approver Comments
add_field_xml '<Field DisplayName="Approver Comments" Name="ApproverComments" Type="Note" NumLines="6" />'

echo "Ensure default view columns"
for fld in Title RequestType Department Status Requestor Manager Modified; do
  m365 spo list view field add --webUrl "$SITE_URL" --listTitle "$LIST_TITLE" --viewTitle "All Items" --fieldTitle "$fld" >/dev/null 2>&1 || true
done

echo "Done. List '$LIST_TITLE' is configured at $SITE_URL"


