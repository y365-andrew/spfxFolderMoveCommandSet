## y-365-folder-move

This extension adds 'Shift' button to the toolbar of document libraries in SharePoint which allows you to move entire folders around whilst retention policies are switched on.

The utility is fairly crude in the way that it opperates and there are plenty of additional performance improvements that could be made. 

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp serve
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Compiling the webpart
Just as with any SPFX webpart simply run the following commands from inside the project directory

```bash
gulp bundle --ship
gulp package-solution --ship
```

### Deploying the webpart
Upload the generated .sppkg file in the dist/sharepoint directory to your app catalog and then add it to your site.

### To-Do
- Update library/folder picker to use the @pnp/js template 
