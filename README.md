# PVSDiskCleanup
This module can be used to make offline updates to a Citrix Provisioning Services VHD file based on entries contained in CSV file. The module will prompt the user for the vDisk to be used, mount the VHD and then, as required, load the SYSTEM and/or SOFTWARE registry hives and perform the updates specified in the separate CSV file.

Before mounting the VHD file the module will use the Citrix McliPSSnapin to determine the PVS farm DB and will subsequently query the DB to determine the vDisk mode (standard or private), the number of devices assignments, and whether or not any PVS devices are actively streaming the vDisk. A summary of the vDisk status is presented to the user and the user is given the option of continuing with the cleanup process, or not.

An example CSV can be found in /PVSDiskCleanup/CSV folder (coming soon)

Note: This is very much a work in progress. In my use the module behaves exactly as expected - but the coding/logic could nevertheless use quite a bit of improvement. Many improvements should be coming soon ...


