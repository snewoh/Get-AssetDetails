#  Powershell Module - AssetDetails
This module will gather details about the asset it is run on.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

[HPWarranty](https://www.powershellgallery.com/packages/HPWarranty/2.6.2) - This is only necessary if you want to fetch the warranty from an HP computer. This can be run during the consolidation of all computers so only needs to be installed on the primary computer.

File Server with a share with modify access to all users. Best to make this hidden with $ at the end of a share name. 
```
\\fileserver-01\audit$\
```

### Installing

You will need to create your own config.xml from config.example.xml provided. This is relatively readable and hopefully not too hard to figure out. I am updating the documentation but slower than I'd like.

## Deployment

To install across computers in a domain, you can use the GPP items provided, by copying the module and files to a server share, and then editing the share location. This should be different to the share to save the asset details, and set to Read Only for non admins. You can copy the xml files and paste them into GPP files & folders in computer preferences.

End with an example of getting some data out of the system or using it for a little demo

## Checking installation

The following command should return a computer object with all the asset details it could find. It will search automatically for a config.xml file in the location of the module. If the config file gives a save location, it will save this to a csv file in that location.
```
Get-AssetDetails
```


