# Object To Excel

## BUILING

1.  Build the project using the command:

    `dotnet build`

1.  Pack the project using the command:

    `dotnet pack /p:Version=<VERSION_NUMBER>`

1.  Navigate to the bin/debug folder and run the following command:

    `dotnet nuget push bin/Debug/<NUGET_PACKAGE> --api-key <API_KEY> --source https://api.nuget.org/v3/index.json`

## INSTALLATION

The library can be downloaded to your project using the folloWing command

    `dotnet add package ObjectToExcel`

## USAGE

## TEST
