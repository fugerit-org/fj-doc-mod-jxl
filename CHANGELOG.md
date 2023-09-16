# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.1.0-sa.1] - 2023-09-16

### Added

- [workflow deploy on branch deploy](.github/workflows/deploy_maven_package.yml)
* [Maven build workflow](.github/workflows/build_maven_package.yml)

### Changed

- exclusions for log4j dependencies (groupId : org.apache.logging.log4j)
- fj-doc version set to 3.0.4

## [1.0.0] - 2023-08-30

### Added

- Added junit test
- Added sonar cloud quality gate configuration
- Added badge for sonar cloud quality gate.
- Added badge for maven repo central version.
- Added badge for license.

### Changed

- Upgrade dependencies: parent fj-doc 1.5.6
- Changelog format changed to [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [0.5.0-rc.001] - 2023-07-02

### Added

- Added CHANGELOG.md

### Changed

- Upgrade dependencies: parent fj-doc 1.1.0-rc.001
- Minimum java required version for compiling : 11
