# MOS Excel 2019 — pilot deployment

1. On a Windows machine with desktop Microsoft Excel and .NET Framework 4.7.2 or later, run `MOS_Excel_2019_Setup_v1.0.0.exe` as administrator.
2. Finish Setup, then launch **MOS Excel 2019** from the Desktop or Start Menu.
3. First validate the installer on 2–3 lab machines before distributing it to the full lab.

The installer places program files and the 21 project packages under Program Files. Training/Testing workbooks in Documents and logs in LocalAppData are user data; uninstall does not remove them. The installer is unsigned, so Windows SmartScreen may show a publisher warning. Keep and compare the published SHA-256 hash when copying the installer between machines.
