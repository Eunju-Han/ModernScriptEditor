# cibctoday-o365-modern-script-editor

## Summary

This web part is similar to the classic Script Editor Web Part and allows you to drop arbitrary script or HTML on a modern page. Also, it has the content link Web Part. Both Web Part functionalities work Audience Targeting with SP Groups and AAD Groups.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.14-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

> No pre-requisites needed at this moment.

## Solution

Solution|Author(s)
--------|---------
cibctoday-o365-modern-script-editor | Eunju Han (eunju1.han@cibc.com, CIBC, [Workplace Profile](https://cibc.workplace.com/profile.php?id=100065805794505&sk=about)) |
contributed content link part | Robert Zhou (pingdu.zhou@cibc.com, CIBC, [Workplace Profile](https://cibc.workplace.com/profile.php?id=100080676725781&sk=about)) |

## Version history

Version|Date|Comments
-------|----|--------
1.3.2|August 15, 2022|Adjust UI to have a div section only when a content link or script exists
1.3.1|August 4, 2022|Fixed scripts to be added under head tag of DOM
1.3.0|July 8, 2022|Add Content Link working with both SP and AAD Groups
1.2.3|June 10, 2022|Add Target Audience with AAD Groups
1.2.2|April 14, 2022|Fixed the pnp sp context issue
1.2.1|April 7, 2022|Redesign the way retrieving SP Groups
1.2.0|March 17, 2022|Add Target Audience with SP Groups
1.1.0|March 8, 2022|Upgrate SPFX 1.13. to 1.4
1.0.0|March 3, 2022|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> If you have spfx-fast-serve and spfx-fast-serve-helpers installed
- in the command-line run:
  - **npm run serve** 

## Deploy
* `gulp clean`
* `gulp build --ship`
* `gulp bundle --ship`
* `gulp package-solution --ship`
* Deploy the `.sppkg` file from `sharepoint\solution` to your tenant App Catalog by using the PS script to have unique permissions for the app. Only an authorized AAD group will have access to manage the app.
* If needed, upload the `.sppkg` file from `sharepoint\solution` to your tenant App Catalog manually for testing purposes
	* E.g.: https://&lt;tenant&gt;.sharepoint.com/sites/AppCatalog/AppCatalog  
* Add the web part to a site collection, and test it on a page

## Features

This extension illustrates the following concepts:

- Re-use existing JavaScript solutions on modern pages
- React
- Office UI Fabric
- spfx-fast-serve

> Notice that the spfx-fast-serve will reduce your gulp serve time significantly. Thanks, Robert for finding this plug-in to boost deployment productivity.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [SPFx development environment compatibility](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/compatibility) - As the SharePoint Framework (SPFx) evolves, make sure the various development tools and libraries are up to date
- [SharePoint Framework v1.14 release notes](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/roadmap)
- [Open source - react script editor](https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-script-editor)
- [Open source - react pnp js example](https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-pnp-js-sample)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [SPFx Fast Serve Tool](https://github.com/s-KaiNet/spfx-fast-serve) - A command line utility, which modifies your SharePoint Framework solution, so that it runs continuous serve command as fast as possible

