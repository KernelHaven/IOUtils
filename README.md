# IOUtils

![Build Status](https://jenkins.sse.uni-hildesheim.de/buildStatus/icon?job=KernelHaven_IOUtils)

A utility plugin for [KernelHaven](https://github.com/KernelHaven/KernelHaven).

Utilities for reading an writing Excel workbooks.

## Usage

Place [`IOUtils.jar`](https://jenkins.sse.uni-hildesheim.de/view/KernelHaven/job/KernelHaven_IOUtils/lastSuccessfulBuild/artifact/build/jar/IOUtils.jar) in the plugins folder of KernelHaven.

This plugin will automatically register the utility classes so that Excel workbooks are supported in all places where previously only CSV was.

## Dependencies

This plugin has no additional dependencies other than KernelHaven.

## License

This plugin is licensed under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0.html).

## Used Libraries

The following libraries are used (and bundled in `lib/`) by this plugin:

| Library | Version | License |
|---------|---------|---------|
| [Apache POI](https://poi.apache.org/) | [3.17](https://archive.apache.org/dist/poi/release/bin/poi-bin-3.17-20170915.zip) | [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0.html) |
