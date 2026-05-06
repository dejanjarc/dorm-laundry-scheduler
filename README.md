# Laundry Scheduler Prototype (Google Apps Script)

A Google Apps Script web app for managing laundry slot reservations in a dorm.

The common issue in student dorms is having a fair system for booking laundry slots that everybody knows how to use and is simple to implement for the majority of dorm representatives and administrators. This application was made exactly because of this: I needed a simple laundry booking system that is easy to hand over and can be used by almost everyone with basic knowledge of Excel, or more specifically, Google Sheets. 

This version is the current working prototype for **Dorm 14**. It includes:

- resident sign in
- self-service account creation
- self-service access code reset with recovery key
- resident dashboard for claiming, moving, and cancelling slots
- admin dashboard with resident management
- spreadsheet-backed data storage
- public board generation
- automation for cycle opening, carryover finalization, and archiving

## Current scope of the application

This version is currently:

- tied to the **Dorm 14** setup (see [here](#current-dorm-14-setup)) with many hardcoded Dorm 14 specific values (e.g. room validation and naming)
- primarily **English-only** 

It works as a prototype, but it is not yet fully generalized for reuse by other dorms.

See [this](#suggested-improvements-todo) part of planned improvements.

---

## Project structure

At the moment, the project is a bit of a mess and has grown into a large monolith. This is to be taken into account for such a prototype. In general, the code is split into 2 main scopes: code for the regular User side and code for the Admin side. 

Below is a short breakdown of the file content.

- #### Server-side files

- `Code.gs`  
  Main app logic, schema setup, reservations, cycles, public board, automation, utilities, resident auth.

- `Admin.gs`  
  Admin dashboard auth, admin sessions, resident management actions.

- #### Client-side files

- `Index.html`  
  Resident sign-in page and resident dashboard UI.

- `AdminPage.html`  
  Admin sign-in page and admin dashboard UI.



---

## Spreadsheet sheets used by the app

The app creates and uses these sheets within the Google Sheets file:

- `Settings`
- `Residents`
- `Slots`
- `Reservations`
- `ReservationsArchive`
- `PublicBoard`
- `Audit`
- `AuditArchive`
- `ResidentsArchive`

---

## Requirements

The point of this project was to make the requirements to deploy this app as simple as possible. 

All a user needs to launch their deployment is:

- **a Google account** (can be a personal account),
- **a new Google Sheets spreadsheet**, which acts as a database,
- **a Google Apps Script project** (connected to the spreadsheet) to have the **project code**.

That's it.

---

## Setup instructions

## 1. Create the spreadsheet

Create a new Google Sheet spreadsheet. This will be the backend database for the app.

Example name:

`Laundry Scheduler Prototype`

---

## 2. Open Apps Script

From the spreadsheet:

- go to **Extensions**
- click **Apps Script**

This project should stay **bound to the spreadsheet**.

This step should take you to a new browser tab, where Google Apps Script opens.

---

## 3. Add the project files

You are now in the Google Apps Script project for your Google Sheets spreadsheet.

Add these files in Apps Script:

- `Code.gs`
- `Admin.gs`
- `Index.html`
- `AdminPage.html`

For each file, **paste in the code** from the files in this Github repository.

**Save** after pasting the code.

---

## 4. Reload the spreadsheet

Go back to the spreadsheet and refresh the browser tab.

A custom menu should appear:

`Dorm 14 Laundry`

You should be able to see it next to other menu options on top of your spreadsheet. 
If it does not appear, refresh again. 

---

## 5. Initialize the project

From the spreadsheet menu, run:

`Dorm 14 Laundry -> 1. Initialize project`

This will:

- store the spreadsheet ID in script properties
- create the required sheets
- write default settings
- prepare the prototype schema

---

## 6. Optional: verify or upgrade schema

If you pasted updated code into an existing prototype, run:

`Dorm 14 Laundry -> 2. Upgrade existing sheet to current schema`

This verifies sheet structure and refreshes dependent data.

> **Note**: you should do this also every time you change the schema of the project in any way to keep the spreadsheet updated.

---

## 7. Configure admin access

Open the `Settings` sheet and review these values:

- `admin_enabled`
- `admin_pin_hash`
- `admin_session_ttl_minutes`
- `admin_allowed_rooms`
- `admin_allowed_resident_ids`

### Important

Admin access is **denied by default**.

A resident only gets the **Admin dashboard** button if:

- `admin_enabled` is `TRUE`
- their room is listed in `admin_allowed_rooms`,
- or their `resident_id` is listed in `admin_allowed_resident_ids`

**Be careful which users get admin access!** 
The user still needs to know the admin PIN to access the admin dashboard, but you should still be careful with giving admin permissions.

### Set the admin PIN

Admin PIN is used to enter the Admin dashboard for any admin users. 

From the spreadsheet menu, run:

`Dorm 14 Laundry -> 14. Set admin PIN`

This stores the hashed admin PIN in settings.

---

## 8. Create residents

You have two options.

### Option A: let residents self-register

Keep these settings enabled in `Settings`:

- `self_signup_enabled = TRUE`
- `self_reset_enabled = TRUE`

Residents can then register from the public web app.

### Option B: seed demo residents

You can do this for testing to make some fake users.

Run:

`Dorm 14 Laundry -> 3. Seed demo residents`

Then run:

`Dorm 14 Laundry -> 4. Generate codes for blank residents`

Access codes are generated into the `Residents` sheet.

---

## 9. Open the first cycle

Before the app is useful, you need to generate the active 2-week cycle.

Run:

`Dorm 14 Laundry -> 5. Open next 2-week cycle now`

This creates slots for the active cycle and opens booking.

---

## 10. Install automation triggers

Run:

`Dorm 14 Laundry -> 10. Install automation triggers`

This installs a time-based trigger for:

- opening the next cycle when due
- finalizing the carryover display of slots from previous cycle
- archiving old data (reservations and audit)

If you ever need to remove them, run:

`Dorm 14 Laundry -> 11. Remove automation triggers`

---

## 11. Deploy as web app

In Apps Script:

- click **Deploy**
- choose **New deployment**
- choose **Web app**
- set **Execute as**: `Me`
- set **Who has access**: `Anyone with the link`  
  or another access level that matches your use case

Deploy and copy the web app URL.

> **Hint**: to help residents remember the link, use an URL shortener service such as *Tinyurl*.

---

## 12. Test resident flow

Open the deployed web app URL. The production (deployed) URL has **/exec** appended at the end. If you wish, you can avoid constant deployements of new versions and instead use the **dev URL**, which you can find under **Test deployment** (the URL is appended by **/dev** at the end).

Test the following actions:

- sign up
- sign in
- claim a slot
- move a slot
- cancel a slot
- save a recovery key
- reset access code

---

## 13. Test admin flow

Make sure the resident account you use for admin is allowed by:

- `admin_allowed_rooms`
- or `admin_allowed_resident_ids`

Then:

- sign in as that resident in the user dashboard
- confirm the **Admin dashboard** button appears
- open the admin dashboard
- sign in with:
  - resident room
  - resident access code
  - admin PIN

Test the following actions:

- updating resident profile
- resetting resident access code
- setting/clearing recovery key
- cancelling resident reservations
- deactivating/reactivating residents
- deleting residents

---

## Main settings to review

Most of the settings in the `Settings` sheet have a description next to them to help you understand their purpose. 

The most important fields are the following:

- `booking_open`
- `active_cycle_start`
- `public_display_start`
- `cycle_days`
- `slot_capacity`
- `timezone`
- `self_signup_enabled`
- `self_reset_enabled`
- `max_active_residents_per_room`
- `resident_stale_after_days`
- `resident_disable_after_days`
- `disabled_account_delete_after_days`
- `activity_touch_cooldown_hours`
- `audit_enabled`
- `dorm_name`
- `archive_keep_past_days`
- `audit_keep_rows`
- `opening_time`
- `simulation_open_at`
- `carryover_finalize_at`

Admin-specific settings:

- `admin_enabled`
- `admin_pin_hash`
- `admin_session_ttl_minutes`
- `admin_allowed_rooms`
- `admin_allowed_resident_ids`

---

## Recommended handover checklist

Before handing the project to another dorm representative:

- verify the deployed web app URL
- verify the `Settings` sheet
- verify admin allowlists
- verify admin PIN is set
- verify triggers are installed
- verify at least one cycle is open
- verify the public board sheet updates correctly
- tell the next owner to redeploy under their own account if ownership changes

---

## Troubleshooting

- ### Custom menu does not appear

    Refresh the spreadsheet page again after saving the script.

- ### Initialization fails

    Make sure the script is bound to the spreadsheet and you are running the menu action from that spreadsheet.

- ### Web app shows old behavior

    After changing Apps Script or HTML:

    - save all files
    - create a new deployment or update the existing deployment
    - reopen the deployment URL

- ### Admin button does not appear

    Check:

    - `admin_enabled = TRUE`
    - the resident is included in `admin_allowed_rooms` or `admin_allowed_resident_ids`

- ### Admin sign-in fails

    Check:

    - the resident can sign into the user dashboard normally
    - the admin PIN was set using `Set admin PIN`
    - the resident is allowed for admin access

- ### Slots or reservations look stale

    Use:

    - `Rebuild cycle availability`
    - `Refresh public board now`
    - `Archive old data now`

    and then refresh the web app.

- ### Automation stopped working

    Reinstall triggers with:

    `Dorm 14 Laundry -> 10. Install automation triggers`

    If the project ownership changes, the new owner should redeploy the web app and recreate the triggers.

---

## Current Dorm 14 setup

As previously mentioned, this version still contains Dorm 14-specific logic. This should change in the future, but for now be aware of these [constraints](#prototype-legacy).

### Hardcoded values in the code

- app name and menu name
- some UI titles and labels
- room validation rules in `isAllowedRoom_()`
- some dorm-specific wording in HTML

This means that if you want to reuse this for another dorm, these changes will need to be implemented:

- move branding into `Settings`
- make room validation configurable
- support multiple UI languages
- remove hardcoded "Dorm 14" text

---

## Prototype legacy

This prototype has some work left to finalize:

- hardcoded dorm-specific text
- hardcoded room validation
- duplicated CSS and client-side UI logic
- large server-side files
- some debug and utility code that should be reviewed
- no full configuration layer for reuse across multiple dorms
- no Slovenian UI yet

---

## Suggested improvements (TODO)

0. Anything left from the [prototype legacy](#suggested-next-improvements-todo) section
1. Move dorm branding into `Settings`
2. Make room validation configurable
3. Add Slovenian and English UI text support
4. Split server-side logic into smaller `.gs` files
5. Split front-end HTML into reusable partials
6. Remove dead/debug code after a cleanup pass

---

## License / ownership

Copyright 2026 Dejan Jarc

Licensed under the Apache License, Version 2.0.
See the LICENSE file for details.