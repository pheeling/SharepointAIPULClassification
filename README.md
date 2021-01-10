# SharepointAIPULClassification
Used to initially classify certain files within certain Document libraries.
Normally this needs a Cloud app Security license and every user within a tenant needs to be licensed.
This script can be used as a "one-time only job" or migration script between two tenants.

# Getting Started
1.	Installation process

    Sync repository to your computer via GIT and checkout all requirements for running the powershell script.
    Run the main.ps1 script from the main.ps1 folder location.

2.	Software dependencies
    - Module AIPService
    - Module ExchangeOnlineManagement
    - AIP Client Installation (https://www.microsoft.com/en-us/download/details.aspx?id=53018)

3.	Latest releases
    Checkout master branch on https://github.com/pheeling/SharepointAIPULClassification/
    Docker version available:
    https://hub.docker.com/repository/docker/pheeeling/aipimage

4.	References

# Build and Test
TODO: Describe and show how to build your code and run the tests.
TODO: Link Docker repo and github code.

# Contribute / Example
This is an example how to interface between different API's and how to use Partner Center to integration into different platforms. Clone this example and create your own repo.

# Important
AIP commands can't be executed against Sharepoint Online files. The script will download each specified file into local directory .\ressources.
Docker version includes a filter function based on the name of the files.
