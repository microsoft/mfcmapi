# Inspecting rules

Select a folder in the message store, then use **Other tables > Rules table**
to inspect its standard rules.
Rules are evaluated in increasing sequence order; rules with the same sequence
have an undefined relative evaluation order.

Client-side metadata may be stored in an associated message in the Inbox. In
the Inbox's associated contents, look for message class `IPM.RuleOrganizer`
and subject `Outlook Rules Organizer`. Its rule stream is client-specific and
is available for raw inspection, not decoded as a portable rules backup.

For a deferred-action message, Smart View can inspect the documented framing of
`PR_CLIENT_ACTIONS`, including multiple concatenated rule actions. Deferred
action payloads remain opaque; inspecting them does not process the message.

If you can locate the deferred-action folder through your provider, inspect
deferred-error messages there for `PR_RULE_ERROR`, `PR_RULE_ACTION_TYPE`,
`PR_RULE_ACTION_NUMBER`, `PR_RULE_PROVIDER`, `PR_RULE_FOLDER_ENTRYID`, and
`PR_RULE_ID`. An action type of zero means the error is not specific to an
action. The action number is zero-based, but zero can also mean the error is
not specific to an action. The folder EntryID and rule ID can help identify the
standard rule; extended rules are not used in deferred-error messages.
Properties can be missing or unavailable. MFCMAPI does not automatically locate
this folder or process or delete these messages.

MrMAPI can inspect standard rules and list extended-rule messages in a folder:

```text
MrMAPI -Rules -Profile "Profile name" -Folder "Folder path"
MrMAPI -RulesJson "rules.json" -Profile "Profile name" -Folder "Folder path"
MrMAPI -HiddenContents -Profile "Profile name" -Folder "Folder path" -MessageClass IPM.ExtendedRule.Message -List
```

`-RulesJson` writes a UTF-8 JSON inspection report containing standard-rule
IDs, names, providers, sequences, and states. Each value names its property and
typed tag. A `status` code appears only when a property could not be read
cleanly, and the state includes decoded flags. It does not include conditions,
actions, extended rules,
or client-specific data. `complete` describes table enumeration, not whether
every property was available. A table error or warning is recorded in
`tableStatus`, which appears only then, makes `complete` false, and causes a
nonzero exit; output-write
errors also cause a nonzero exit. The output path is overwritten if it exists.
Invalid UTF-16 in a reported value stops report generation rather than
replacing the invalid text.
Protect the report as sensitive mailbox data; it is not an `.rwz` export or a
restorable backup.

For the Inbox organizer message, use `-HiddenContents -MessageClass
IPM.RuleOrganizer -Subject "Outlook Rules Organizer" -List`. The filters require
provider support. Remove `-List` and specify `-Output` to export messages;
protect those files as sensitive mailbox data. These are inspection exports,
not Outlook `.rwz` exports or guaranteed restorable backups. Check diagnostic
output for provider and table errors; an exit code alone does not establish that
the results are complete.

These views depend on provider support and access. A failed, cancelled, or
limited table load can be incomplete. They do not represent every client-side
rule; client-specific data may be opaque. Rule properties and exported data can
contain sensitive mailbox information.

See [MS-OXORULE: Extended Rules Message Syntax](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/99ba591c-a741-4b10-afee-729a2b849ef2).
See also [MS-OXORULE: Creating Rich Client-Side Rules](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/950a82ff-acde-40b3-8f2f-fc46ceb7a182).
See [MS-OXORULE: PidTagClientActions Property](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/81b16a56-c4db-42a4-8e51-10e38bf055a6).
See [MS-OXORULE: PidTagRuleState Property](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/a0d934eb-4b76-49ee-a342-536aa64afcfd).
See [MS-OXORULE: PidTagRuleSequence Property](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/0f04b9ec-59ad-431d-820f-d4be9bff0cc8).
See [MS-OXORULE: DEM Syntax](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/dd4df994-bd8c-43ae-9556-80676e8e5ca0) and [Processing a DEM](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/e092bd63-766c-444f-badf-08bb69e18b07).
See [MS-OXORULE: PidTagRuleActionNumber Property](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/954192e2-2b0a-4dc9-9975-908dd9b95daf).
