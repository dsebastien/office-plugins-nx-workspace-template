# OfficePlugins

This project provides a Nrwl NX workspace containing a Word plugin based on Angular 10+, as well as instructions (below) to side-load the plugins in Office apps.

This was inspired by the following article: https://medium.com/@ragavanrajan/building-office-add-in-using-angular-8-209624ba61ed

## Installing
Run `npm install`
Done.

## Starting the Word plugin
Run `npm run start`

Or more generally: `npx nx serve:word`

Once started, you can access it locally through http://localhost:4200/index.html

Note that once you depend on Office-specific APIs (e.g., read the content of a Word document), then it'll become harder/impossible to test your app like that.

## Side-loading the Word plugin
On Windows (only option), you need to create a Windows share on the folder containing the manifest.xml file:

* Open Windows Explorer
* Go to <project folder>/apps/word/src
* Right click on the assets folder and click on "Properties"
* Go to the "Sharing" tab
* Click on "Share"
* Share with yourself
* Take note of the share path, aka UNC (e.g., \\<hostname>\Users\<username>\whatever\office-plugins-nx-workspace-template\apps\word\src\assets)

Once done, the manifest.xml file is now available via a Windows Share.

Now, to side-load the plugin, open up Word and:

* Go to File > Options > Trust Center
* Click on "Trust Center Settings"
* Click on "Trusted Add-in Catalogs"
* Add the UNC to the "Catalog Url" field and click on "Add catalog"
* Once added, make sure to check the checkbox in the "Show in Menu" column below 
* Once done, click on Ok until everything is saved
* Close Word

Now, reopen word. You should find "My Add-ins" in the "Insert" tab. If not, then you failed one of the steps.

If the button is there then:
* Click on it
* Click on "SHARED FOLDER"
* You should see the plugin. Click on it

At this point the add-in should be loaded and you should see a button for it on the top right corner.

To refresh the plugin after changes:
* Go back to the "SHARED FOLDER" tab in "My Add-ins"
* Click refresh
* Click on the add-in again

If the plugin fails to load with the following error "We can't open this add-in from localhost", then:
* Open a command prompt as Administrator
* Run the following command: `CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"` 
* Reference: https://docs.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost

To test against Word online:
* Go to the "Insert" tab
* Click on "Office Add-ins"
* Click on "Upload my add-in"
* Browse for the manifest.xml file

Reference explanations:
* https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins
* https://www.youtube.com/watch?v=XXsAw2UUiQo

## Quick Start & Documentation
This project was generated using [Nx](https://nx.dev).

[Nx Documentation](https://nx.dev/angular)

[10-minute video showing all Nx features](https://nx.dev/angular/getting-started/what-is-nx)

[Interactive Tutorial](https://nx.dev/angular/tutorial/01-create-application)

## Adding capabilities to your workspace

Nx supports many plugins which add capabilities for developing different types of applications and different tools.

These capabilities include generating applications, libraries, etc as well as the devtools to test, and build projects as well.

Below are our core plugins:

- [Angular](https://angular.io)
  - `ng add @nrwl/angular`
- [React](https://reactjs.org)
  - `ng add @nrwl/react`
- Web (no framework frontends)
  - `ng add @nrwl/web`
- [Nest](https://nestjs.com)
  - `ng add @nrwl/nest`
- [Express](https://expressjs.com)
  - `ng add @nrwl/express`
- [Node](https://nodejs.org)
  - `ng add @nrwl/node`

There are also many [community plugins](https://nx.dev/nx-community) you could add.

## Generate an application

Run `ng g @nrwl/angular:app my-app` to generate an application.

> You can use any of the plugins above to generate applications as well.

When using Nx, you can create multiple applications and libraries in the same workspace.

## Generate a library

Run `ng g @nrwl/angular:lib my-lib` to generate a library.

> You can also use any of the plugins above to generate libraries as well.

Libraries are sharable across libraries and applications. They can be imported from `@office-plugins-nx-workspace-template/mylib`.

## Development server

Run `ng serve my-app` for a dev server. Navigate to http://localhost:4200/. The app will automatically reload if you change any of the source files.

## Code scaffolding

Run `ng g component my-component --project=my-app` to generate a new component.

## Build

Run `ng build my-app` to build the project. The build artifacts will be stored in the `dist/` directory. Use the `--prod` flag for a production build.

## Running unit tests

Run `ng test my-app` to execute the unit tests via [Jest](https://jestjs.io).

Run `nx affected:test` to execute the unit tests affected by a change.

## Running end-to-end tests

Run `ng e2e my-app` to execute the end-to-end tests via [Cypress](https://www.cypress.io).

Run `nx affected:e2e` to execute the end-to-end tests affected by a change.

## Understand your workspace

Run `nx dep-graph` to see a diagram of the dependencies of your projects.

## Further help

Visit the [Nx Documentation](https://nx.dev/angular) to learn more.

## ☁ Nx Cloud

### Computation Memoization in the Cloud

<p align="center"><img src="https://raw.githubusercontent.com/nrwl/nx/master/images/nx-cloud-card.png"></p>

Nx Cloud pairs with Nx in order to enable you to build and test code more rapidly, by up to 10 times. Even teams that are new to Nx can connect to Nx Cloud and start saving time instantly.

Teams using Nx gain the advantage of building full-stack applications with their preferred framework alongside Nx’s advanced code generation and project dependency graph, plus a unified experience for both frontend and backend developers.

Visit [Nx Cloud](https://nx.app/) to learn more.
