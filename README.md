# New Inventory Tool V2

New Inventory Tool V2 is a Windows desktop application for looking up inventory,
checking the live state of computers and completing device-rounding work. It is a
PowerShell/WPF application: the PowerShell script contains the application logic,
the XAML file defines the interface, and the CSV files under `Data` provide the
inventory and lookup data.

> **Intended environment:** this tool is designed for authorized technicians on
> the organization network. Network, WMI/CIM, DNS, CMDB, and rounding-page
> features only work when the current Windows account is permitted to access the
> relevant systems.

## Features

### Inventory search and device summary

- Search the loaded inventory by device name/host name, asset tag, or serial
  number. Barcode-scanner input works like keyboard input and can be submitted
  with **Enter**.
- Recognize computers, monitors, carts, scanners, and microphones, and resolve a
  peripheral back to its parent computer. Tangent cart relationships are also
  followed so cart peripherals appear under the appropriate computer.
- Display type, host name, asset tag, serial number, parent, RITM, retirement
  date, last-round date, and device location.
- Show associated parent, child, and cart-child devices in one table, including
  direct CMDB links when an asset has a supported link.
- Double-click an associated device to make it the active summary item.
- Detect supported peripheral names that do not match their parent and offer a
  **Fix Name** action. Association/name changes are recorded in
  `Output/CMDBUpdates.csv` for follow-up; the source inventory exports are not
  rewritten.
- Add a peripheral by scanning/searching its asset tag and remove selected
  peripheral associations. These changes are likewise queued in
  `Output/CMDBUpdates.csv`.
- Validate the peripherals physically reported by a reachable computer against
  the inventory. Built-in laptop displays and unusable placeholder monitor
  serials are excluded from monitor validation.

### Network and live diagnostics

- Automatically perform a short connectivity check after a successful query.
- Display online/offline state, response time, IPv4 address, named subnet,
  uptime, and pending-reboot state when available.
- Open an explicit, continuous `ping -t` command window with **Ping**.
- Retrieve **Live Details** through remote Windows management, including system,
  operating-system, hardware, storage, memory, network, power/battery/dock,
  monitor, user-profile, signed-in-user, last-input, uptime, and reboot
  information.
- Resolve an arbitrary IPv4 address against `Data/SiteSubnets.csv` with
  **Lookup Subnet**. `10.64.x.x` addresses are identified as VPN addresses.
- Show a large, easy-to-read monitor label containing the parent asset tag and
  host name.

### Location editing

- View city, location, building, floor, room, and department for the selected
  device.
- Edit location through cascading choices built from the site's computer export,
  location master data, and previously saved user additions.
- Preserve a newly entered valid location relationship in the site's
  `LocationMaster-UserAdds - <site>.csv` file.
- Stage the selected device's location change until its rounding event is saved,
  preventing another queried device from inheriting an unsaved edit.

### Rounding workflow

- Choose a maintenance type: **Mobile Cart**, **General Rounding**,
  **Critical Clinical**, or **Excluded**.
- Record check status, time spent, comments, cable management, monitor labelling,
  peripheral operation, physical cart operation, cabling requirements, and
  whether the device must be added to the tracker.
- Use **Check Complete** to toggle all applicable completion checks at once.
- Let the rounding timer count from the three-minute default, or adjust it with
  the `+` and `-` controls.
- Save parent-computer rounding records to `Output/RoundingEvents.csv`. Queries
  for peripherals are resolved to their parent computer before the event is
  written.
- Open a configured web rounding page with **Manual Round** for supported parent
  computers. The button explains why it is unavailable when the device type,
  parent, mapping, or URL is unsuitable.
- Track daily and weekly progress in the footer. Click **Days/week** to select
  planned rounding dates, and click **Data File** to inspect the age/status of
  loaded source files.

### Nearby workflow

- After an event is saved, build a list of computers at the same location. More
  saved locations can be accumulated into the active Nearby scope.
- Review host, IP address, subnet, asset/location fields, maintenance type,
  last-round date, days since rounding, and editable status.
- Filter the table by rounded today, excluded, recently rounded, and critical
  clinical devices, or use **Show All**.
- Sort by the displayed columns, resize columns, and scroll horizontally with a
  touchpad or **Shift + mouse wheel**.
- Select one or more rows and right-click to **Isolate**, ping the selected
  hosts, or apply an inaccessible status in bulk. Double-click a row to return to
  its full System view.
- Use **Ping All** to test every currently visible row and populate reachable
  network details with progress shown in the footer.
- Save all rows with a selected status to `Output/RoundingEvents.csv` using
  **Save**, or discard the accumulated location scopes with **Clear List**.

### Event-file editor and updates

- Open `RoundingEvents.csv` from **File Editor**, edit cells with constrained
  dropdowns where appropriate, and delete one or several rows. Blank required
  values are highlighted; unsaved changes prompt before the window closes.
- Refresh the Nearby table and progress badges after edited event data is saved.
- Update the application from the repository's `main` branch with the included
  updater. The updater can create a dated full backup, preserves site
  `LocationMaster*.csv` files and `Output` data when updating without a full
  backup, and recreates the desktop shortcut.

## Requirements

- Windows with Windows PowerShell 5.1, WPF, and .NET Framework available.
- Permission to run local PowerShell scripts. The supplied launcher uses
  `-ExecutionPolicy Bypass` for this process only.
- Read access to the inventory CSV files in `Data` and write access to the
  application folder (especially `Output` and site location-master files).
- Network/DNS access for ping and subnet information.
- Appropriate remote WMI/CIM/DCOM permissions for validation, Live Details,
  uptime, and reboot status.
- Internet access is required only for the updater, CMDB links, and configured
  manual-rounding webpages.

## Files and data

Keep these files together; the launcher and application resolve paths relative
to their installation directory.

| Path | Purpose |
| --- | --- |
| `Launch-NewAssetTool-matching-v14.bat` | Normal application launcher. |
| `NewAssetTool.Wpf.matching.v14.ps1` | Application logic. |
| `NewAssetTool.matching.v14.xaml` | Window layout and visual styles. |
| `Data/<Site>/Computers - <Site>.csv` | Computer inventory and authoritative location relationships. |
| `Data/<Site>/Monitors - <Site>.csv` | Monitor inventory. |
| `Data/<Site>/LocationMaster*.csv` | Location choices and technician-added relationships. |
| `Data/Carts.csv`, `Data/Scanners.csv`, `Data/Mics.csv` | Shared peripheral inventory. |
| `Data/SiteSubnets.csv` | CIDR-to-site/subnet lookup table. |
| `Data/Rounding.csv` | Asset-to-rounding-page mapping. |
| `Output/RoundingEvents.csv` | Created/updated rounding-event log. |
| `Output/CMDBUpdates.csv` | Created/updated queue of association and naming changes. |

`Output` is created automatically if it does not exist. Do not manually change
the headers in source or output CSV files. If direct event correction is needed,
prefer the built-in **File Editor** so the current schema and permitted values are
retained.

## Installation and launch

1. Place the entire `New-Inventory-Tool-V2` folder on the Windows computer. The
   included updater installs it to the current user's Desktop by default.
2. Confirm that the appropriate site folders and current CSV exports exist under
   `Data`.
3. Double-click **`Launch-NewAssetTool-matching-v14.bat`**.
4. If Windows displays a security warning, verify that the files came from the
   expected repository/administrator before allowing them to run.
5. At first launch, choose the applicable inventory site if prompted. The app
   loads that site's computer and monitor exports plus the shared peripheral
   files.

For troubleshooting from a command prompt, launch the same application directly:

```bat
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File ".\NewAssetTool.Wpf.matching.v14.ps1" -XamlPath ".\NewAssetTool.matching.v14.xaml"
```

The `-STA` argument is required by WPF.

## Instruction manual

### 1. Find and inspect a device

1. Enter or scan a host name, asset tag, or serial number in the search box.
2. Press **Enter** or select **Query**.
3. Review the System header and summary. A connectivity check runs shortly after
   the inventory match; successful checks reveal network and system-status data.
4. Review **Associated Devices**. Double-click a row when you want the summary
   card to describe that parent or peripheral.
5. If **Fix Name** appears, confirm the selected peripheral and parent are
   correct, then select it to stage the expected name in `CMDBUpdates.csv`.

Clearing the search box clears the selected System device but does not erase the
current Nearby list.

### 2. Use network tools

- **Ping:** after querying a device, select **Ping** to open a separate continuous
  ping window. Close that command window when finished. The automatic post-query
  check is a single short check and does not open a command window.
- **Live Details:** wait for a successful connectivity check to enable the
  button, then select it. Retrieval can take a few seconds. Failure normally
  means the host is offline, DNS is unavailable, or the current account lacks
  remote-management access.
- **Monitor Label:** select this after a query to display the parent computer's
  asset tag and host name at large size for labelling work.
- **Lookup Subnet:** enter a valid IPv4 address in the dialog, select **Lookup**,
  and read the matching name. An unknown result means no matching CIDR is present
  in `Data/SiteSubnets.csv`.

### 3. Add, remove, or validate peripherals

1. Query the computer or one of its associated devices.
2. To add an association, select **Add Peripheral**, scan or enter the
   peripheral's asset tag, inspect the preview, and confirm **Add**.
3. To remove an association, select the relevant associated-device row(s), then
   select **Remove Peripheral** and confirm the requested operation.
4. Select **Validate** to compare discovered external monitors/peripherals with
   inventory. The parent computer must be online and remotely accessible.
5. Treat `Output/CMDBUpdates.csv` as a work queue: adding, removing, or fixing a
   name does not directly update the upstream CMDB or inventory exports.

### 4. Review or change the location

1. Check all six location fields: city, location, building, floor, room, and
   department.
2. Select **Edit Location**.
3. Work from the top down through the cascading selectors. Later choices are
   refreshed from earlier choices; department is intentionally available from
   the site's complete department list.
4. Select **Save** in the location panel, or **Cancel** to discard the edit.
5. Complete and save the rounding event. The location edit is committed with
   that event; querying a different device first discards the staged edit.

The main **Save Event** control remains unavailable when required location data
is blank. Hover over a disabled control for the specific reason when available.

### 5. Record a System rounding event

1. Query the device. If it is a peripheral, verify that the displayed parent is
   the computer that should receive the event.
2. Choose the correct **Maintenance Type** and **Check Status**.
3. Complete the applicable checks individually, or select **Check Complete** to
   mark all currently applicable checks and set the status to Complete.
4. Adjust the recorded minutes if necessary and enter useful comments.
5. Mark **Cabling needed** and/or **Add device to device tracker** when follow-up
   is required.
6. If a supported rounding webpage is required, select **Manual Round** before
   saving. A successful launch is reflected in the saved event; merely saving an
   ordinary event leaves it pending for the external rounding processor.
7. Select **Save Event**. The app writes the parent-computer record, adds its
   displayed location to Nearby scope, refreshes Nearby, resets the form, and
   returns focus to the search box for the next scan.

An **Excluded** maintenance type is deliberately not saved through the normal
System event flow. Use the applicable excluded workflow/status rather than
recording it as an ordinary completed round.

### 6. Process Nearby devices

1. Save at least one System event to establish a location scope, then open the
   **Nearby** tab. Additional saved System events can add more locations.
2. Use the filter checkboxes to include or hide rounded-today, excluded,
   recently-rounded, and critical-clinical rows. Select **Show All** to reset an
   isolation and expose the full scoped list.
3. To focus on a subset, select rows with **Ctrl**/**Shift**, right-click, and
   choose **Isolate**. Right-click also offers **Ping selected host(s)** and bulk
   inaccessible statuses.
4. Use **Ping All** to check all visible rows. Do not close the app until the
   footer reports that the batch has completed.
5. Set each row's status in the Status column. Rows whose current rounding rules
   make them read-only display their recorded status as text.
6. Select **Save**. Only rows with a chosen status are written; the app reports
   how many events were saved. These entries are marked internally as originating
   from Nearby so the view can immediately reflect today's work.
7. Double-click a row to inspect it on the System tab. When returning after a
   save, the app restores the prior Nearby view where possible.
8. Select **Rebuild Nearby** to reload the scoped rows from current inventory and
   event data, or **Clear List** to remove all active scopes.

### 7. Correct event records

1. On the Nearby tab, select **File Editor**.
2. Edit a cell directly. Check status, yes/no fields, and maintenance type use
   dropdowns to prevent unsupported values.
3. Delete one row with its red **X**, or select multiple rows, right-click, and
   choose **Delete selected row(s)**. Deletion is applied only when saved.
4. Select **Save & Close**. If you close the window with unsaved changes, choose
   whether to save, discard, or cancel closing.

The editor changes `Output/RoundingEvents.csv` directly. Make corrections
carefully, because saved edits and deletions affect progress badges, last-rounded
dates, and Nearby status calculations.

### 8. Set the rounding plan and read footer status

- Select the **Days/week** badge, choose at least one planned rounding date, and
  confirm. The app uses the plan with `RoundingEvents.csv` to calculate today's,
  this week's, and remaining-per-day counts.
- Select the **Data File** badge to review which inventory inputs were loaded and
  their file ages.
- The center status badge reports ready, found/not found, working, saved, ping
  progress, completion, and warning states.
- The footer displays the active `Data` and `Output` paths. Check these before
  editing source files or collecting output for another process.

## Updating

1. Close the application.
2. Double-click **`Update-New-Inventory-Tool-V2.bat`** while connected to the
   internet.
3. If an existing Desktop installation is found, choose whether to create the
   recommended timestamped full backup.
4. Wait for the updater to download and extract `main`, restore protected user
   data, and recreate **New Inventory Tool V2.lnk** on the Desktop.
5. Confirm the final **Done** message before closing the window.

Without a full backup, the updater preserves only `Data/**/LocationMaster*.csv`
and the contents of `Output`; other local modifications in the installed folder
are removed. For diagnostic logging, run the updater from Command Prompt with
`--log`. Its log is written to `Desktop\Update-New-Inventory-Tool-V2.log`.

## Troubleshooting

| Symptom | What to check |
| --- | --- |
| The app does not open | Keep the `.bat`, `.ps1`, and `.xaml` together; use the supplied launcher so PowerShell runs in STA mode. Check `NewAssetTool.startup-error.log` beside the script. |
| Device not found | Confirm the correct site was selected, refresh the site's CSV exports, and try the host name, asset tag, or serial without extra spaces. |
| Online device shows offline | Confirm DNS/network reachability and that ICMP echo is permitted. DNS alone is not treated as proof that a host is online. |
| Live Details or validation fails | Confirm the target is online and the account/firewall permits remote WMI/CIM/DCOM access. |
| Subnet is Unknown | Confirm the IP is IPv4 and that `Data/SiteSubnets.csv` contains a matching CIDR entry. |
| Save Event is disabled | Query a valid device, resolve all required location fields, and check the control's tooltip. Events must resolve to a parent computer. |
| Manual Round is disabled | Query a supported parent-computer family and verify that its asset has a valid entry/URL in `Data/Rounding.csv`. |
| No Nearby rows appear | Save a System event to establish a location scope, clear restrictive filters, then select **Rebuild Nearby** or **Show All**. |
| Changes seem missing from inventory | Association and name fixes are intentionally queued in `Output/CMDBUpdates.csv`; they do not rewrite the exported inventory. |
| Horizontal grid content is hidden | Drag column dividers/the bottom scrollbar, use a two-finger horizontal gesture, or hold **Shift** while using the mouse wheel. |

## Development and tests

The automated test suite performs static regression checks against the
PowerShell and XAML implementation. From a development environment with Python
and `pytest` installed, run:

```bash
python -m pytest -q
```

The production application itself must still be exercised on Windows because a
non-Windows test environment cannot launch the WPF interface or validate the
organization's remote-management dependencies.
