# NamedSolutionExplorer

Allows creation of new Solution Explorer windows with actual names

## Project Structure
Refer to the readme.md in each project for a more complete description

### NamedSolutionExplorerViewerService
This service is the primary service responsible for interacting with menu items and coordinating with the IDE

### NewSolutionExplorerViewer
This is the command-handler for the new menuitem on the solution explorer

### NewSolutionexplorerViewerPackage
This is the VSIX package class for the extension

### Models\NSESettings
The load/save configuration for a NamedSolutionExplorer window

### Repositories\SettingsRepository
Respnsible for storing the settings in-memory as well as loading/saving them from/to disk




