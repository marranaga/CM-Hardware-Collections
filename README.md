# CM Hardware Collections

Generate unique device collections for each manufacturer and model of devices in your SCCM instance

## Settings

To set this up for your environment, edit the `Settings.json` file for your environment.

### Prefix

Key type: `[String]`

If you would like each device collection to start with a specific prefix, you can add it here. Standard naming convention is `[Prefix-]Manufacturer[-Model]`

### RootFolderPath

Key type: `[String]`

Specifies the root folder path to save all new device collections in SCCM, starting from `$SiteCode\DeviceCollection`

### Computers

Key type: `[Boolean]`

Specifies whether or not to create unique Device Collections for each make and model of computer

### VideoCards

Key type: `[Boolean]`

Specifies whether or not to create unique Device Collections for each make and model of video card