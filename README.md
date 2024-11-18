K2 TypeScript Broker Template
===

This template demonstrates best-practices for developing a K2 Broker using TypeScript.

<!-- TOC -->
* [K2 TypeScript Broker Template](#k2-typescript-broker-template)
  * [Features](#features)
  * [Getting Started](#getting-started)
  * [Running Unit Tests](#running-unit-tests)
  * [Building your bundled JS](#building-your-bundled-js)
  * [Creating a service type](#creating-a-service-type)
  * [License](#license)
  * [Migration from Ava to Vitest](#migration-from-ava-to-vitest)
<!-- TOC -->

## Features

- Full object model intellisense for making development easier
- Sample broker code that accesses jsonplaceholder.
- Sample unit tests with mocks and code coverage.
- Vite configuration for TypeScript.

## Getting Started

This template requires [Node.js](https://nodejs.org/) v22.5.1+ to run.

Install the dependencies and devDependencies:

```bash
npm install
```

Alternatively, use pnpm to reduce the size of the `node_modules` directory:

```bash
npm install -g pnpm # Only required once
pnpm install
```

See the documentation for [@k2oss/k2-broker-core](https://www.npmjs.com/package/@k2oss/k2-broker-core)
for more information about how to use the broker SDK package.

## Running Unit Tests

To run the unit tests, run:

```bash
npm test
pnpm test # Alternative
```

You will find the code coverage results in [coverage/index.html](./coverage/index.html).

## Building your bundled JS

When you're ready to build your broker, run the following command

```bash
npm run build
pnpm run build # Alternative
```

You will find the results in the [dist/index.js](./dist/index.js).

## Creating a service type

Once you have a bundled .js file, upload it to your repository (anonymously
accessible) and register the service type using the system SmartObject located
at System > Management > SmartObjects > SmartObjects > JavaScript Service
Provider and run the Create From URL method.

## License

MIT, found in the [LICENSE](./LICENSE) file.

[www.k2.com](https://www.k2.com)

## Migration from Ava to Vitest
If you used a previous version of this template and want to migrate to Vitest, please review the [Migration.md](./Migration.md) guide.