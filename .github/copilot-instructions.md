# Copilot Instructions for Java365

## Repository Overview

Java365 is a Java library for interacting with **SharePoint** list data via the legacy SOAP (`Lists.asmx`) web service. It supports **SharePoint Online** (Office 365) and on-premises SharePoint deployments.

## Project Structure

```
Java365/
├── connector/               # Main source package (package connector)
│   ├── _Constants.java      # Shared URLs and charset constants
│   ├── SharePointClient.java# Claims-based (SAML) authentication
│   ├── Manager.java         # List operations: read, insert, attach files
│   └── ListsRequest.java    # CAML XML batch request builder
├── wsdl/
│   └── wsdlfile.wsdl        # SharePoint Lists web service WSDL
├── src/
│   └── test/java/connector/ # JUnit 5 test classes
│       ├── ConstantsTest.java
│       ├── ListsRequestTest.java
│       ├── ManagerXmlTest.java
│       └── ManagerAuthTest.java
├── pom.xml                  # Maven build descriptor
└── .gitignore
```

> **Important:** The `connector/` directory sits at the repository root, **not** under `src/main/java`. Maven is configured with `<sourceDirectory>${project.basedir}</sourceDirectory>` to handle this non-standard layout.

## Build and Test

**Requirements:** JDK 17+, Maven 3.9+, and `wsimport` (from JDK 8, placed earlier in `PATH` than JDK 17). `wsimport` was removed in JDK 9+ so a JDK 8 installation is required solely to provide this tool at build time; the runtime target remains Java 17.

```bash
# Compile sources (generates SOAP stubs from WSDL, then compiles everything)
mvn compile

# Run all 56 unit tests (fully offline — no SharePoint tenant needed)
mvn test

# Compile, test, and package into a JAR
mvn package

# Skip tests for a faster build
mvn package -DskipTests

# Run a single test class
mvn test -Dtest=ManagerXmlTest

# Clean build
mvn clean test
```

The compiled JAR is placed at `target/java365-1.0-SNAPSHOT.jar`.

## Testing Conventions

- All tests use **JUnit 5 (Jupiter)** and live under `src/test/java/connector/`.
- Tests are fully **offline** — no SharePoint connection required.
- SOAP interactions are mocked with **Mockito**.
- Test class naming: `<SourceClass>Test.java` or `<SourceClass><Feature>Test.java`.
- New tests should be added to the existing test classes where appropriate, or as a new `*Test.java` class in `src/test/java/connector/`.

## Coding Conventions

- **Java 17** language level (LTS). Java 21 is also compatible.
- Source encoding: **UTF-8**.
- Package name: `connector` (matches the `connector/` directory at the root).
- No external formatting configuration; follow the existing code style in the `connector/` files.
- Avoid adding new dependencies unless strictly necessary. Current dependencies: JAX-WS RI (`com.sun.xml.ws:rt`), JUnit 5, Mockito.
- Do **not** commit `src/META-INF/SAML.xml` — it contains credentials and is excluded via `.gitignore`.

## Key Architecture Notes

- **WSDL-based stubs**: The build runs `wsimport` against `wsdl/wsdlfile.wsdl` and generates SharePoint SOAP client stubs into `target/generated-sources/wsimport/com/microsoft/schemas/sharepoint/soap/` at build time. Do not edit these generated files.
- **Authentication**: `SharePointClient` performs SAML/claims-based authentication to obtain `FedAuth` and `rtFa` cookies. `Manager.sharePointListsAuth()` creates an authenticated SOAP port using those cookies.
- **Security**: `Manager.generateXmlNode()` blocks XXE/DOCTYPE attacks by using a hardened `DocumentBuilderFactory`. Do not relax these security settings.
- **On-premises support**: `Manager.createManagerService()` and `SharePointClient` have overloaded constructors that accept an explicit protocol (e.g., `"http://"`) or a custom STS URL (e.g., ADFS endpoint). NTLM/Basic auth is supported by passing `null` as the `cookieToken`.
