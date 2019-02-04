# Backup NetScaler

## Requirements

* PowerShell Version 4.0
* Administrator Privileges
* NetScaler Module from PSGallery (Please update to latest Version, 1.7.0 as of writing)
* at least Windows Server 2012
* Network Connectivity between NetScaler Appliance and Windows Server on
  * HTTP/S
  * SSH

## How to use

* Download resources
* verify requirements
* run Backup-NetScaler_GUI.ps
* Verify created Backup Job

## Known Error Database

| Error Code | Description | Resolution |
| ---------- | ----------- | ---------- |
| 599 | Too many backups. NetScaler cannot maintain more than 50 backups. | Check Backup count in web interface and remove until count is below 50 |
