# SharePoint Online React Template 

## Currently on Version 3.1.0

**View the [Changelog](./CHANGELOG.md) for more Information.**

<img src=./docs/readme_banner.png width="1200" />

## Overview 📖

- [Prerequisites 🛠](#prerequisites-)
  - [Used SPFx Version](#used-spfx-version)
  - [Environment](#environment)
  - [NPM Packages](#npm-packages)
- [Cloning this Project 🚀](#cloning-this-project-)
- [TailwindCSS](#tailwindcss)
- [Settings 🛠](#settings-)
  - [Project Name](#project-name)
  - [Development](#development)
- [Deployment 🛫](#deployment-)
  - [Test](#minimal-path-to-awesome-test-deployment)
  - [Production](#maximum-path-to-awesome-production-deployment)
  - [Upload to App Catalog](#how-to-upload-to-app-catalog)
  - [Configure Solution Settings](#configure-solution-settings)
- [Add App to Page](#how-to-add-to-the-sharepoint-online)
- [Step-By-Step Guide 🏄](#step-by-step-guide-)
- [Optional 💡](#optional-)
  - [Full-Width Column](#full-width-column)
  - [Webpart Icon](#webpart-icon)
- [Development](#development)
  - [Component Libary 📚](#component-library-)
  - [Package Solution Assets](#package-solution-assets-lists)
- [Property Pane 📋](#property-pane-)
- [SPFx Lifecycle 🔄](#spfx-lifecycle-hooks-)
  - [Webpart](#webpart)
  - [Property Pane](#property-pane-1)
- [REST Api 📥](#rest-api-)
- [Logging 📃](#logging-)
- [Upgrade SPFx Version](#upgrade-spfx-version)
- [Git Commits ✉](#git-commits-)
- [Known Issues 🚧](#known-issues-)

<br />

## Prerequisites 🛠

First check out the official recommended development environment setup for SharePoint SPFx: [Link](https://learn.microsoft.com/de-de/sharepoint/dev/spfx/set-up-your-development-environment)

If the Versions in that Documentation has been changed here are the Versions this Template got created with:

### Used SPFx Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

See [here](https://learn.microsoft.com/de-de/sharepoint/dev/spfx/roadmap) for the official Microsoft SharePoint SPFx Changelog.

### Environment

| Software | Version |
| -------- | ------- |
| Node     | v22.17.1 |
| npm      | 10.9.2   |
| SPFx      | 1.21.1  |

To use the correct Node version install [nvm](https://github.com/nvm-sh/nvm). It makes changing the Node version much easier!

Run the following in the terminal:

```bash
nvm use #this will take the specified version in the .nvmrc file

# optional
nvm install
```

For more detailed Information check out the [nvm usage documentation](https://github.com/nvm-sh/nvm#usage).

### Environment

Check out the detailed Guide from Microsoft to get the latest versions:

https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment

## Cloning this Project 🚀

If you are on Github just create a new Repository and select the "sharepoint-online-react-template" as a Template.

If you are cloning this Repository follow this Guide:

**First of all run the following command to clone this project:**

Replace `<your_username>` with your git username or instead copy the Repository URL at the top right.

If you want to create the Repository in a different folder name, just replace `<project_name>` with your Project Name.

```bash
git clone https://git.snutig.dev/<your_username>/SharePoint-Online-React-Template.git <project_name>
```

<img src=./docs/clone_repo.png width="900" />

<br />

**Change local Repository URL:**

After you cloned this Repository and want to configure it to your own needs, you have to change the associated Git Repository. Otherwise your commits would be pushed to this Template Repository in theory (but you dont have permisisons for this ... ).

To change the Repository Remote URL run the following command:

_Replace `git://new.url.here` with your new created Repository URL_

```bash
git remote set-url origin git://new.url.here

git push -u origin main
```

**Important 🚨: Install NPM Packages:**

Run the following code in your command line. This will take some time.

```bash
npm install
```

## TailwindCSS

> TailwindCSS is activated by default at Version 4.1.11 since 12.08.2025!

See here for more Information about TailwindCSS: https://tailwindcss.com/

See this post to check out how we are getting the Fluent UI / SharePoint Theme Variables working:

https://github.com/tailwindlabs/tailwindcss/discussions/15199

```css
...
@theme inline {
    /* Colors coming from your Fluent/brand CSS variables */
  --color-neutral-foreground: var(--colorNeutralForeground1);
  --color-neutral-foreground-on-brand: var(--colorNeutralForegroundOnBrand);
  --color-brand-background: var(--colorBrandBackground);
  --color-brand-background-hover: var(--colorBrandBackgroundHover);
}
```

For more Color Variables check out the [FluentUI Documentation](https://react.fluentui.dev/?path=/docs/theme-colors--docs).

If the Tailwind CLI is not working on macOS try installing it with homebrew:

```bash
brew install tailwindcss
```

## SPFx Fast Serve Tool

See official Package description [here](https://github.com/s-KaiNet/spfx-fast-serve?tab=readme-ov-file).

Best Tool ever for SharePoint Development!!!

## Settings 🛠

## Project Name 🔨

If you have choosen a Name for your Webpart you have to change some variables in the whole Code.

Lets say you have choosen the Name `BestWebpart`. In this example we will go through every Setting you have to make to apply the Name properly.

**`package.json`**

```json
{
  "name": "best-webpart", // change the name here (use kebab case!)
  ...
}
```

**Webpart Files `src/webparts/`**

Change all the Occurrences of `WebpartName` to your choosen Webpart Name (in this example `ClothingShop` / `ClothingShopWebPart`).

If your VSCode is properly configured it should ask you to change all the imports or do it automatically. Otherwise you have to adjust them by yourself (no tutorial for this, sorry ... ).

<img src=./docs/file_name_change.png width="400" />

**`src/webparts/ClothingShop/ClothingShopWebPart.manifest.json`**

```json
{
  ...
  "id": "<new uuid>", // create a new uuid
  ...
  "alias": "ClothingShopWebPart", // change the name here (use kamel case!)
  ...
}
```

You can create a new uuid here: [here](https://www.uuidgenerator.net/version4)

> ignore the warning produced from the comments in this file. You can also delete the comments to remove the Error Messages.

**`config/config.json`**

```json
{
  ...
  "bundles": {
    "clothing-shop": { // you can choose your own name here
      "components": [
        {
          "entrypoint": "./lib/webparts/ClothingShop/ClothingShopWebPart.js",
          "manifest": "./src/webparts/ClothingShop/ClothingShopWebPart.manifest.json"
        }
      ]
    }
  },
  ...
  "localizedResources": {
    "ClothingShopWebPartStrings": "lib/webparts/ClothingShop/loc/{locale}.js"
  }
}
```

**`src/webparts/ClothingShop/ClothingShopWebPart.ts`**

```typescript
...
export interface IClothingShopWebPart {
  ...
}

export default class ClothingShopWebPart extends BaseClientSideWebPart<IClothingShopWebPart> {
  ...
}
```

### Deployment

**`config/package-solution.json`**

```json
{
  ...
  "solution": {
    ...
    "id": "<new uuid>",
    ...
  }
}
```

You can create a new uuid here: [here](https://www.uuidgenerator.net/version4)

### Development 🧑‍💻

**`config/serve.json`**

```json
{
  ...
  "initialPage": "https://enter-your-SharePoint-site/_layouts/workbench.aspx" // change the URL here and replace "enter-your-SharePoint-site" with your SharePoint URL
}
```

## Deployment 🛫

This Section describes how to deploy your Webpart to your SharePoint Tenant. Feel free to copy this step to your documentation because its the same for every Webpart.

_Requires that you have cloned this repository and made the previous changes._

### Minimal Path to Awesome (Test Deployment) 🔬

Open a new Terminal in the root folder of this Project and run the following Code:

```bash
gulp trust-dev-cert # only for first deployment

gulp serve # will open a new tab in your default browser with the workbench
```

### Maximum Path to Awesome (Production Deployment) 💡

```bash
gulp clean # optional

gulp bundle --ship

gulp package-solution --ship
```

This commands will create a new `.sppkg` file in the `sharepoint/solution` folder. This file contains all the information and code for your Webpart.

### How to upload to App Catalog

Navigate to the App Catalog in your SharePoint Tenant. Usually the URL looks like this:

`https://<tenant>.sharepoint.com/sites/appcatalog/SitePages/Home.aspx`

The Page looks like this:

<img src=./docs/app_catalog_1.png width="850" />

Click on the first square/button. On the next page just drag and drop the previously created `.sppkg` file in the table.

**Upload with NodeJS**

> This Section hasnt been tested yet! 🚧

First you have to authenticate yourself in the Azure Active Directory. After you got the necessary Credentials follow the Guide [here](https://github.com/pnp/pnpjs/blob/version-3/docs/getting-started.md#using-pnpsp-spfi-factory-interface-in-nodejs).

After you got the Authentication its possible to upload the `.sppkg` with the ALM api provided by the [PnP sp](https://pnp.github.io/pnpjs/sp/alm/#pnpspappcatalog) package.

Run the following commands in the command line:

```bash
openssl req -x509 -newkey rsa:2048 -keyout keytmp.pem -out cert.pem -days 365 -passout pass:yourPasswordHere

openssl rsa -in keytmp.pem -out key.pem -passin pass:yourPasswordHere
```

> The above Code will generate three files, `cert.pem`, `key.pem`, and `keytmp.pem`. Upload the `cert.pem` file to Azure AD application registration. The `key.pem` file will be read in the configuration.

```typescript
// this represents the file bytes of the app package file
const blob = new Blob();

// there is an optional third argument to control overwriting existing files
const r = await catalog.add("myapp.app", blob);

// this is at its core a file add operation so you have access to the response data as well
// as a File instance representing the created file
console.log(JSON.stringify(r.data, null, 4));

// all file operations are available
const nameData = await r.file.select("Name")();
```

**Upload with Powershell**

> This Section hasnt been tested yet! 🚧

This method uses the [PnP PowerShell](https://pnp.github.io/powershell/) package.

```Powershell
#Parameters
$AppCatalogURL = "https://tenant-name.sharepoint.com/sites/apps"
$AppFilePath = "./sharepoint/solution/sharepoint-app.sppkg"

#Connect to SharePoint Online App Catalog site
Connect-PnPOnline -Url $AppCatalogURL -Interactive

#Add App to App catalog - upload app to sharepoint online app catalog using powershell
$App = Add-PnPApp -Path $AppFilePath

#Deploy App to the Tenant
Publish-PnPApp -Identity $App.ID -Scope Tenant
```

### Configure Solution Settings

Solution Settings file is saved in `config/package-solution.json`.

Follow [this Guide](https://pnp.github.io/blog/post/spfx-21-professional-solutions-superb-solution-packages/) to configure all the mandatory settings.

## How to add to the SharePoint Online:

### 1. Click on Neu and choose Seite:

<img src=./docs/deploy1.png width="900" />

### 2. Click on Apps and choose your App:

<img src=./docs/deploy2.png width="900" />

## Step-By-Step Guide 🏄

This Guide assumes that you have copied all files from this repository to your local environment and have all [Prerequisites](#prerequisites) installed.

### 1. Install Dependencies

```bash
nvm use 16 #make sure to use Node Version 16

npm install
```

### 2. Trust Dev Certificate

```bash
gulp trust-dev-cert # necessary for SPFX Development
```

### 3. Change Workbench URL

`config/serve.json`

```json
{
  ...
  "initialPage": "https://<tenant>.sharepoint.com/layouts/workbench.aspx"
  // example: "https://testcompany.sharepoint.com/sites/Test/_layouts/workbench.aspx"
}
```

### 4. Change Webpart ID's

Create for **each ID** a new UUIDv4. This can be done [here](https://www.uuidgenerator.net/version4).

`config/package-solution.json`

```json
{
  ...
  "solution": {
    ...
    "id": "<new uuid1>"
  }
}
```

`src/webparts/webpartName/WebpartNameWebPart.manifest.json`

```json
{
  "id": "<new uuid2>"
}
```

If this Step is skipped the Webpart cant be deployed to the final Page because every Webpart ID has to be unique.

### 5. Run Development Mode

Now you should be able to run the Workbench and actually see your Webpart.

```bash
gulp serve

# press ctrl + c to exit
```

### 6. Change Configuration

`package.json`

```json
{
  "name": "<project name>",
  ...
}
```

`config/package-solution.json`

```json
{
  ...
  "solution": {
    "name": "<project name>",
    ...
    "metadata": {
      "shortDescription": {
        "default": ""
      },
      "longDescription": {
        "default": ""
      }
    }
  },
  "paths": {
    "zippedPackage": "solution/<project-name>.sppkg"
  }
}
```

**For all Configurations see the Instructions [here](#project-name).**

## Optional 💡

### Full-Width Column

To have the Webpart use the full width of the Website its mandatory to adjust some settings and enable the Full-Width Column. But there are some Restrictions you have to consider.

**Create new Web**

Be patient when you create a new Web. The full width column is **not available on Team Sites**. We have to create a new **Communication Site**.

<img src=./docs/full_width_column_web.png width="850" />

<br />

**`manifest.json`**

```json
{
  ...
  "supportsFullBleed": true,
  ...
}
```

Now its possible to select the Webpart in a full width column on a SharePoint Site.

To add it follow the instructions on this screenshot:

<img src=./docs/full_width_column_section.png width="400" />

<br />

**Development for Full-Width Column**

In the Workbench the Webpart will always be just in a small part of the page. To make it fullscreen use the following styling in your `WebpartName.module.scss` file. This will make the webpart use the full size of the screen.

```css
:global {
  #workbenchPageContent,
  .CanvasComponent.LCS .CanvasZone {
    max-width: 100% !important;
  }
}
```

### Webpart Icon

[Official Documentation](https://learn.microsoft.com/de-de/sharepoint/dev/spfx/web-parts/basics/configure-web-part-icon)

To customize your Webpart Icon edit the following line in the `WebpartNameWebPart.manifest.json` file:

**You can only use one of the following options!**

```json
{
  ...
  "preconfiguredEntries": [
    {
      ...
      // Option 1 (Image URL):
      "iconImageUrl": "https://example.website.com/image.png",
      // Option 2 (Fluent UI Icon):
      "officeFabricIconFontName": "Sunny",
      // Option 3 (Base64 Image):
      "iconImageUrl": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAB4AAAAQ4CAIAAABnsVYUAAFP2klEQVR42uzd..."
    }
  ]
}
```

You can find all the available Fluent UI Icons [here](https://developer.microsoft.com/en-us/fluentui#/styles/web/icons).

## Development

### Component Library 📚

**Fluent UI**

Fluent UI is developed by the Microsoft Team and integrates its components perfectly into the SharePoint/Microsoft environment.
Its also already installed in the dependencies of this project (`@microsoft/sp-office-ui-fabric-core`) so you can use their components out of the box.
The possibility to work with [themes](https://github.com/microsoft/fluentui/wiki/theming) in Fluent UI is also pretty useful and makes styling easier.

**Useful Resources:**

- [Fluent UI Components](https://developer.microsoft.com/en-us/fluentui#/controls/web)
- [Fluent UI Icons](https://developer.microsoft.com/en-us/fluentui#/styles/web/icons)

### Package Solution Assets (Lists)

In the Package Solution Assets its possible to put models for Lists that will be created when the App is installed.

*Important: Works only for Site-App-Installation, not for global Installation!*

With this Feature its possible to update Lists with new Features and that is pretty useful.

**Useful Resources:**

 - [Assets Guide](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/provision-sp-assets-from-package)
 - [Field Schema](https://learn.microsoft.com/en-us/sharepoint/dev/schema/field-element-field?redirectedfrom=MSDN)
 - [List Schemas](https://learn.microsoft.com/en-us/sharepoint/dev/schema/list-instances)
 - [List Schema](https://learn.microsoft.com/en-us/sharepoint/dev/schema/list-element-list)
 - [List Template ID's](https://learn.microsoft.com/de-de/archive/blogs/vinitt/list-of-feature-id-listtemplate)

## Property Pane 📋

**Useful Resources:**

- [PnP SPFx Controls](https://pnp.github.io/sp-dev-fx-property-controls/)
- [SPFx Property Pane Configuration](https://learn.microsoft.com/de-de/sharepoint/dev/spfx/web-parts/basics/integrate-with-property-pane)

## SPFx Lifecycle Hooks 🔄

### Webpart

Quoted from [this site](https://www.vrdmn.com/2019/12/sharepoint-framework-web-part-and.html).

When loading the web part on a page the methods are executed in the following order:

```typescript
// 1.
protected onAfterDeserialize(deserializedObject: any, dataVersion: Version): TProperties;
// 2.
protected onInit(): Promise<void>;
// 3.
protected render(): void;
// 4.
protected onBeforeSerialize(): void
```

When the web part is removed from a page the methods are executed in the following order:

```typescript
// 1.
protected onDispose(): void;
```

### Property Pane

**Open**

```typescript
// 1.
protected loadPropertyPaneResources(): Promise<void>
// 2.
protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
// 3.
protected onPropertyPaneRendered(): void;
// 4.
protected onPropertyPaneConfigurationStart(): void;
```

**Update**

There are two modes that have to be differentiated: reactive mode or non-reactive mode.

> Reactive implies that changes made in the PropertyPane are transmitted to the web part instantly and the user can see instant updates. This helps the page creator get instant feedback and decide if they should keep the new configuration changes or not. NonReactive implies that the configuration changes are transmitted to the web part only after "Apply" PropertyPane button is clicked."

reactive mode

```typescript
// 1.
protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void;
// 2.
protected render(): void;
// 3.
protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
// 4.
protected onPropertyPaneRendered(): void;
// 5.
protected onPropertyPaneConfigurationComplete(): void;
```

non-reactive mode

```typescript
// 1.
protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void;
// 2.
protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
// 3.
protected onPropertyPaneRendered(): void;
```

Events after the "Apply" Button got clicked:

```typescript
// 1.
protected onAfterPropertyPaneChangesApplied(): void;
// 2.
protected render(): void;
// 3.
protected onPropertyPaneConfigurationComplete(): void;
// 4.
protected onPropertyPaneRendered(): void;
```

**Close**

```typescript
// 1.
protected onPropertyPaneConfigurationComplete(): void;
```

## REST Api 📥

Working with the REST Api in SPFx Webparts is pretty easy. The best solution is to use the `@pnp/sp` package because it already provides you with all the important functions in a promise-based system.

Insert the following code to your `WebpartNameWebPart.ts` file:

```typescript
import { spfi, SPFx } from "@pnp/sp";
...
protected async onInit(): Promise<void> {
  await super.onInit();
  const sp = spfi().using(SPFx(this.context));
  ...
}
...
```

## Logging 📃

When it comes to logging in your application there is an out of the box tool by Microsoft which you can use, but it works a little bit different than the native `console.log()`. Its still possible to use the native `console.log()` script but there are cases when the Microsoft Logging is better or more useful.

To use the Microsoft Logging Tool follow this example:

```typescript
import { Log } from '@microsoft/sp-core-library';

...
const LOG_SOURCE: string = "HelloWorldWebPart";

Log.info(LOG_SOURCE, 'Hello World this is an info log.')
Log.verbose(LOG_SOURCE, 'Hello, this is a verbose log.')
Log.warn(LOG_SOURCE, 'Hello, this is a waning log.')
Log.error(LOG_SOURCE, new Error('Hello, this is a error log!')) //Log.error needs an Error Object
```

Check out the [official Documentation](https://learn.microsoft.com/en-us/javascript/api/sp-core-library/log?view=sp-typescript-latest).

> The Log class provides static methods for logging messages at different levels (verbose, info, warning, error) and with context information. Context information helps identify which component generated the messages and allows for filtering of log events. In a SharePoint Framework application, these messages will appear on the developer dashboard.

In the Log method description its written that the logs are shown in the _developer dashboard_. So what is that? I also didnt heard of this before but after some research i found a Microsoft documentation [here](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-developer-dashboard).

To open the _developer dashbord_ in SharePoint you have to press the keys <kbd>Ctrl</kbd> + <kbd>F12</kbd> or on MacOS also works <kbd>Command</kbd> + <kbd>F12</kbd> .

## Upgrade SPFx Version

You can use the PnP Cli to upgrade the version with the Assistant. Look here for further information: [Link](https://pnp.github.io/cli-microsoft365/cmd/spfx/project/project-upgrade/)

```bash
m365 spfx project upgrade [options]

# example
m365 spfx project upgrade md > "upgrade-report.md"
```

## Git Commits ✉

We are using the packages `husky` and `commitlint` to keep control of the git commit messages. The configuration can be found in the `commitlint.config.js` file.

`commitlint` will throw an error if the commit message is not like the expected syntax or the "type" is not existing in the list of avaliable types. You can find the available types in the `commitlint.config.js` file. [Here](https://www.conventionalcommits.org/en/v1.0.0/#summary) you can find some explanations for the types but most of them are self explaining.

An error can look like this:

<img src=./docs/commitlint_error.png width="900" />

## Known Issues 🚧

**`module.scss` not found**

After you create a new `module.scss` file it sometimes isnt recognized in the code when you try to import the styles into your `.tsx` file.

<img src=./docs/issue_scss_not_found.png width="600" />

To solve this error just run `gulp serve` in the command line.

**Error - [webpack] 'dist': [object Object]**

Replace the code in `node_modules/@microsoft/gulp-core-build-webpack/lib/WebpackTask.js` at line 84 with the following code:

```ts
this.logError(`'${outputDir}':` + os_1.EOL + statsResult.errors.map(x => JSON.stringify(x, null, 2)).join(os_1.EOL) + os_1.EOL);
```