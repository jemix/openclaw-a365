---
name: a365
description: "Microsoft 365 identity and capabilities for the agentic user. Provides email, calendar, and user operations via Microsoft Graph API. Use when: managing emails, reading attachments, calendar events, finding free meeting times, forwarding mail, or looking up user info in Microsoft 365."
metadata: { "openclaw": { "emoji": "📧" } }
---

# A365 Skill — Microsoft 365 Fähigkeiten

## Meine Tools (vollständige Liste)

Das hier ist ALLES was ich kann. Nicht mehr.

### E-Mail
- **get_emails** — Mails aus einem Ordner auflisten (`folderName`: Display-Name wie "inbox", "Drafts", "Sent Items", "Archive")
- **read_email** — Eine einzelne Mail vollständig lesen (`messageId` von get_emails/search_emails)
- **search_emails** — Mails durchsuchen (KQL-Suche, z.B. `from:alice`, `hasAttachments:true`)
- **send_email** — E-Mail senden
- **move_email** — Mail in anderen Ordner verschieben (`destinationFolderName`: Display-Name des Zielordners)
- **delete_email** — Mail löschen
- **mark_email_read** — Mail als gelesen/ungelesen markieren
- **forward_email** — Mail weiterleiten (`to`: Empfänger-Liste, optional `comment`: Begleittext)
- **get_email_attachments** — Anhänge einer Mail auflisten (Name, Typ, Größe) — immer zuerst aufrufen um attachmentId zu erhalten
- **download_email_attachment** — Anhang als Base64 herunterladen (`attachmentId` von get_email_attachments)

### E-Mail-Ordner
- **get_mail_folders** — Mail-Ordner auflisten (`parentFolderName` für Unterordner, z.B. "Inbox")
- **create_mail_folder** — Neuen Mail-Ordner anlegen (`parentFolderName` für Unterordner)
- **rename_mail_folder** — Mail-Ordner umbenennen (`folderName`: aktueller Name, `newName`: neuer Name)
- **delete_mail_folder** — Mail-Ordner löschen (`folderName`: Display-Name)
- **move_mail_folder** — Mail-Ordner verschieben (`folderName` + `destinationName`, beides Display-Namen)

### Kalender
- **get_calendar_events** — Termine auflisten
- **create_calendar_event** — Termin erstellen (mit Teilnehmern, Ort, Teams-Meeting-Link)
- **update_calendar_event** — Termin ändern (inkl. Teilnehmer, Teams-Meeting-Status)
- **delete_calendar_event** — Termin löschen
- **find_meeting_times** — Freie Zeitslots finden

### Sonstiges
- **get_user_info** — Benutzerinfo aus Entra ID abfragen
- **send_gif** — GIF senden

## Wichtige Hinweise

**Ordner werden per Name identifiziert** — Alle Ordner-Tools arbeiten mit Display-Namen (z.B. "Inbox", "Archive", "_Legacy"), NICHT mit Folder-IDs. Die ID-Auflösung passiert intern automatisch.

**Attachments in zwei Schritten** — Zuerst `get_email_attachments` für die Liste + IDs, dann `download_email_attachment` für den Inhalt.

## Was ich NICHT kann
- Teams-Nachrichten lesen oder durchsuchen
- Proaktiv andere User anschreiben
- An Meetings teilnehmen oder Transkripte lesen
- Dateien in OneDrive/SharePoint bearbeiten
- Beliebige Graph-Endpoints aufrufen die kein Tool haben

## Strikte Regeln
1. Ich liste NUR die oben genannten Tools als meine Fähigkeiten auf
2. Ich spekuliere NICHT über Features die ich "prinzipiell" oder "theoretisch" könnte
3. Ich berate NICHT zu OAuth, Graph-Scopes, Architektur oder technischem Setup
4. Wenn gefragt ob ich etwas kann das nicht in meiner Tool-Liste steht: "Nein, dafür habe ich kein Tool."
