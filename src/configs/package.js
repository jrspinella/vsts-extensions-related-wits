var exec = require("child_process").exec;

// Package release extension
var command = `tfx extension create --overrides-file configs/release.json --manifest-globs vss-extension.json --rev-version --no-prompt --json`;
exec(command, { cwd: 'dist/' }, (error, stdout) => {
    if (error) {
        console.error(`Could not create package: '${error}'`);
        return;
    }

    console.log(`Package created`);
});

// Package dev extension
var manifest = require("../vss-extension.json");
var extensionId = manifest.id;
var command = `tfx extension create --overrides-file configs/dev.json --manifest-globs vss-extension.json --rev-version --extension-id ${extensionId}-dev --no-prompt --json`;
exec(command, { cwd: 'dist/' }, (error, stdout) => {
    if (error) {
        console.error(`Could not create package: '${error}'`);
        return;
    }

    console.log(`Package created`);
});