# Fugerit Document Generation Framework (fj-doc)

## JXL Renderer (XLS)(fj-doc-mod-jxl)

[back to fj-doc index](https://github.com/fugerit-org/fj-doc.git)  

[![Maven Central](https://img.shields.io/maven-central/v/org.fugerit.java/fj-doc-mod-jxl.svg)](https://mvnrepository.com/artifact/org.fugerit.java/fj-doc-mod-jxl) 
[![license](https://img.shields.io/badge/License-Apache%20License%202.0-teal.svg)](https://opensource.org/licenses/Apache-2.0)
[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=fugerit-org_fj-doc-mod-jxl&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=fugerit-org_fj-doc-mod-jxl)
[![Coverage](https://sonarcloud.io/api/project_badges/measure?project=fugerit-org_fj-doc-mod-jxl&metric=coverage)](https://sonarcloud.io/summary/new_code?id=fugerit-org_fj-doc-mod-jxl)

*Deprecation* (2022-11-17):  
Previously this module was part of the [Fugerit Doc Venus](https://github.com/fugerit-org/fj-doc.git) and now is formally deprecated and put in an separate repository.

*Description* :  
Type handlers for generating documents in XLS format using last version of  
[JExcelApi](https://mvnrepository.com/artifact/net.sourceforge.jexcelapi/jxl/2.6.12).

*Status* :  
All basic features are implemented.  
The code is mature, and as JXL 2.X is deprecated, this module is not likely to be updated.  
Better to use the module based on Apache POI [fj-doc-mod-poi](https://github.com/fugerit-org/fj-doc.git) 
  
  
*Quickstart* :  
Basically this is only a type handler, see core library [fj-doc-base](https://github.com/fugerit-org/fj-doc.git).  
NOTE: If you have any special need you can open a pull request or create your own handler based on this.

See [CHANGELOG.md](CHANGELOG.md) for details.