# Xls Spring Boot Starter

## Description

This project provides an easy-to-use Spring Boot starter for working with xls files into Spring Boot applications, 
enabling seamless Excel file (xlsx) generation and editing.
In particular you will have two main ways of operating:

* Import an xlsx template with all the styles and formatting already set,
  then modify it by inserting the values and finally saving it.
* Create an xlsx file from scratch, populate it with values and styles and save it.

## Features

- **Excel File Creation**: Easily create new Excel files with support for various data types.
- **Excel File Editing**: Edit existing Excel files, including adding or removing rows/columns, updating cell values, and more.

## Getting Started

### Prerequisites

Ensure you have the following installed:
- JDK 17 or later
- Maven 3.6+ or Gradle 7.3+

### Adding the Starter to Your Project

Add the following dependency to your `pom.xml`:

```xml
<dependency>
    <groupId>org.jspring</groupId>
    <artifactId>xls-spring-boot-starter</artifactId>
    <version>0.0.1-SNAPSHOT</version>
</dependency>
```

## Configuration
You can customize some properties in your application.yml (or application.properties if you prefer)
```yaml
spring:
  export:
    xlsx:
      # the path to the xlsx template to modify
      templatePath: src/main/resources/template/Blank.xls 
```

## Basic Usage
On importing the dependency on the classpath you will have some beans autoconfigured 
and you can use anywhere in your spring beans/services the xlsx feautures autowiring the following services:

- XlsxCreateService
- XlsxReadingService
- XlsxWritingService
- XlsxSearchingService
- XlsxTableService
- XlsxCellsWritingService

```java
@Autowired
private XlsxReadingService xlsxReadingService;

public void readXlsFromTemplate() {
  XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
  // ... other stuff
}
```


## Project structure

```
project-root
│
├── .gitignore
├── pom.xml
├── Readme.md
│
└── src
├── main
│   ├── java
│   │   └── org
│   │       └── jspring
│   │           └── xls
│   │               ├── config
│   │               ├── domain
│   │               ├── enums
│   │               ├── service
│   │               └── utils
│   │
│   └── resources
│
└── test

└── target
```

## Contributing
We welcome contributions! Please see our contribution guidelines for details.

## License
This project is licensed under the MIT License - see the LICENSE file for details.