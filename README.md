# Java365 — SharePoint Online SOAP Connector

A Java library for interacting with **SharePoint Online** (Office 365) list data via the legacy SOAP (`Lists.asmx`) web service. It handles claims-based authentication, CAML query construction, list item CRUD, and file attachments.

---

## Table of Contents

1. [Requirements](#requirements)
2. [Project Structure](#project-structure)
3. [Quick Start — Linux / macOS](#quick-start--linux--macos)
4. [Quick Start — Windows 11](#quick-start--windows-11)
5. [Building the Project](#building-the-project)
6. [Running the Tests](#running-the-tests)
7. [Using the Connector in Your Code](#using-the-connector-in-your-code)
8. [Authentication: SAML Template](#authentication-saml-template)
9. [Troubleshooting](#troubleshooting)

---

## Requirements

| Tool | Minimum version | Notes |
|------|-----------------|-------|
| **JDK** | 17 (LTS) | Java 21 (LTS) also works. OpenJDK or any compatible distribution (Temurin, Zulu, Corretto, …). Must include `wsimport` — see note below. |
| **Maven** | 3.9 | Tested with Apache Maven 3.9.x |
| **wsimport** | bundled with JDK 8 | Required only at build time. See [wsimport note](#wsimport-note) below. |
| Internet access (build) | — | Maven downloads dependencies from Maven Central on first build. |
| SharePoint Online tenant | — | Required only at runtime for live calls. Tests run fully offline. |

### wsimport note

`wsimport` was removed from JDK 9+.  
The build locates it via the system `PATH`. Install a JDK 8 alongside your JDK 17 to provide it:

- **Linux (apt)**: `sudo apt-get install temurin-8-jdk` and the tool is at  
  `/usr/lib/jvm/temurin-8-jdk-amd64/bin/wsimport`. Add that directory to your `PATH` **before** the JDK 17 `bin` directory.
- **macOS (Homebrew)**: `brew install --cask temurin@8` installs it at  
  `/Library/Java/JavaVirtualMachines/temurin-8.jdk/Contents/Home/bin/wsimport`. Add that to `PATH`.
- **Windows 11**: download Temurin 8 from [adoptium.net](https://adoptium.net) and add its `bin` folder to the front of `PATH` (see [Windows section](#quick-start--windows-11) for full steps).

---

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

> **Important:** The `connector/` directory sits at the repository root, not under `src/main/java`. This is the original project layout and is preserved intentionally. Maven is configured with `<sourceDirectory>${project.basedir}</sourceDirectory>` to handle it.

---

## Quick Start — Linux / macOS

### 1 — Install a JDK 17 (or 21)

**Debian / Ubuntu:**
```bash
sudo apt-get update
sudo apt-get install -y temurin-17-jdk     # Eclipse Adoptium
java -version                               # should print openjdk 17...
```

**macOS (Homebrew):**
```bash
brew install --cask temurin@17
java -version
```

### 2 — Install JDK 8 (provides `wsimport`)

**Debian / Ubuntu:**
```bash
sudo apt-get install -y temurin-8-jdk
```

**macOS (Homebrew):**
```bash
brew install --cask temurin@8
```

### 3 — Put `wsimport` on your `PATH`

Make sure `wsimport` resolves but `java` / `javac` still point to JDK 17:

**Debian / Ubuntu** — add to `~/.bashrc` (or `~/.zshrc`):
```bash
export PATH="/usr/lib/jvm/temurin-8-jdk-amd64/bin:$PATH"
export JAVA_HOME="/usr/lib/jvm/temurin-17-jdk-amd64"
```
Then reload: `source ~/.bashrc`

**macOS** — add to `~/.zshrc`:
```bash
export PATH="/Library/Java/JavaVirtualMachines/temurin-8.jdk/Contents/Home/bin:$PATH"
export JAVA_HOME="/Library/Java/JavaVirtualMachines/temurin-17.jdk/Contents/Home"
```
Then reload: `source ~/.zshrc`

Verify:
```bash
java -version       # should show 17
wsimport -version   # should show wsimport version "2.x.x"
```

### 4 — Install Maven

**Debian / Ubuntu:**
```bash
sudo apt-get install -y maven
mvn -version
```

**macOS (Homebrew):**
```bash
brew install maven
mvn -version
```

### 5 — Clone and build

```bash
git clone https://github.com/crsuarez/Java365.git
cd Java365
mvn package -DskipTests     # compile + package (skip tests for speed)
```

### 6 — Run the tests

```bash
mvn test
```

Expected output:
```
[INFO] Tests run: 11, Failures: 0, Errors: 0, Skipped: 0  -- ManagerXmlTest
[INFO] Tests run: 10, Failures: 0, Errors: 0, Skipped: 0  -- ConstantsTest
[INFO] Tests run: 14, Failures: 0, Errors: 0, Skipped: 0  -- ListsRequestTest
[INFO] Tests run: 16, Failures: 0, Errors: 0, Skipped: 0  -- ManagerAuthTest
[INFO] Tests run: 51, Failures: 0, Errors: 0, Skipped: 0
[INFO] BUILD SUCCESS
```

---

## Quick Start — Windows 11

### 1 — Install JDK 17 (or 21)

1. Go to [adoptium.net](https://adoptium.net/temurin/releases/?version=17) and download the **Windows x64 `.msi`** installer for Temurin 17.
2. Run the installer. Choose **"Add to PATH"** and **"Set JAVA_HOME"** during setup.
3. Open a new **Command Prompt** or **PowerShell** and verify:
   ```powershell
   java -version
   ```

### 2 — Install JDK 8 (provides `wsimport`)

1. Go to [adoptium.net](https://adoptium.net/temurin/releases/?version=8) and download the **Windows x64 `.msi`** for Temurin 8.
2. Run the installer. **Do not** let it override `JAVA_HOME` or overwrite the existing `PATH` entry for JDK 17 — only install the files.
3. After installation, note the JDK 8 `bin` directory. It is typically:
   ```
   C:\Program Files\Eclipse Adoptium\jdk-8.0.xxx-hotspot\bin
   ```

### 3 — Put `wsimport` on `PATH` (while keeping JDK 17 as default)

Open **System Properties → Environment Variables** (or use PowerShell):

```powershell
# In an elevated PowerShell window:
$jdk8bin = "C:\Program Files\Eclipse Adoptium\jdk-8.0.xxx-hotspot\bin"
[System.Environment]::SetEnvironmentVariable(
    "PATH",
    "$jdk8bin;" + [System.Environment]::GetEnvironmentVariable("PATH", "Machine"),
    "Machine"
)
```

Replace `jdk-8.0.xxx-hotspot` with the exact directory name installed on your machine.

Open a **new** Command Prompt and verify:
```cmd
java -version       # should show 17
wsimport -version   # should show wsimport version "2.x.x"
```

> **Important:** `wsimport` must appear first in `PATH`. If `java -version` now shows 8, move the JDK 17 `bin` entry back to the top of `PATH`.

### 4 — Install Maven

1. Download the latest **Binary zip archive** from [maven.apache.org/download](https://maven.apache.org/download.cgi) (e.g. `apache-maven-3.9.x-bin.zip`).
2. Extract to a directory such as `C:\tools\apache-maven-3.9.x`.
3. Add `C:\tools\apache-maven-3.9.x\bin` to `PATH` (same Environment Variables dialog).
4. In a new Command Prompt:
   ```cmd
   mvn -version
   ```

### 5 — Clone and build

```cmd
git clone https://github.com/crsuarez/Java365.git
cd Java365
mvn package -DskipTests
```

### 6 — Run the tests

```cmd
mvn test
```

---

## Building the Project

The build performs the following steps automatically in order:

1. **`initialize`** — Creates the `target/generated-sources/wsimport` output directory.
2. **`generate-sources`** — Runs `wsimport` against `wsdl/wsdlfile.wsdl` to generate the SharePoint SOAP client stubs into `target/generated-sources/wsimport/com/microsoft/schemas/sharepoint/soap/`.
3. **`compile`** — Compiles the generated stubs together with the `connector/` source files.
4. **`test-compile`** — Compiles the test classes under `src/test/java/`.
5. **`test`** — Runs all JUnit 5 tests via Maven Surefire.

### Common build commands

| Command | What it does |
|---------|-------------|
| `mvn compile` | Generate stubs + compile sources only |
| `mvn test` | Compile + run all 51 unit tests |
| `mvn package -DskipTests` | Compile + package JAR without running tests |
| `mvn clean test` | Full clean rebuild + tests |
| `mvn test -pl . -Dtest=ManagerXmlTest` | Run a single test class |

The compiled JAR is placed at `target/java365-1.0-SNAPSHOT.jar`.

---

## Running the Tests

All tests are fully offline — they do **not** require a SharePoint tenant or any network connection.

```bash
# Run all tests
mvn test

# Run a specific test class
mvn test -Dtest=ListsRequestTest

# Run a specific test method
mvn test -Dtest=ManagerXmlTest#generateXmlNodeBlocksXxeDoctypeDeclaration

# Run tests with verbose output
mvn test -Dsurefire.useFile=false
```

### Test coverage summary

| Test class | # tests | What it covers |
|------------|--------:|----------------|
| `ConstantsTest` | 10 | Every constant in `_Constants` |
| `ListsRequestTest` | 14 | CAML batch XML structure, all three request types, `createListItem` field generation and error paths |
| `ManagerXmlTest` | 11 | `generateXmlNode()` XML parsing, XXE/DOCTYPE attack blocking, `xmlToString()` output format |
| `ManagerAuthTest` | 16 | `sharePointListsAuth` preconditions, null-safety guards, mocked SOAP port delegation |
| **Total** | **51** | |

---

## Using the Connector in Your Code

### Step 1 — Authenticate

```java
// Initialise the endpoint (called once per session, before any other Manager calls)
Manager.createManagerService("yoursite.sharepoint.com", "/sites/your-site");

// Obtain a cookie token via claims-based (SAML) authentication
SharePointClient client = new SharePointClient(
        "user@yoursite.onmicrosoft.com",
        "YourPassword",
        "yoursite.sharepoint.com");

// Build an authenticated SOAP port
ListsSoap port = Manager.sharePointListsAuth(
        "user@yoursite.onmicrosoft.com",
        "YourPassword",
        client.getCookieNedToken());
```

### Step 2 — Read a list

```java
ArrayList<String> columns = new ArrayList<>(List.of("Title", "Description", "Created"));
Manager.displaySharePointList(port, "My List Name", columns, "100");
```

### Step 3 — Insert a list item

```java
HashMap<String, String> fields = new HashMap<>();
fields.put("Title", "New item title");
fields.put("Body",  "Item content goes here");
Manager.insertListItem(port, "My List Name", fields);
```

### Step 4 — Add a file attachment

```java
// Attach a local file to the last inserted item
Manager.addItemWithAttachmentToList(
        port,
        "My List Name",
        fields,
        "/path/to/document.pdf",
        "document.pdf");
```

---

## Authentication: SAML Template

`SharePointClient` reads a SAML XML template at startup from:

```
<project-root>/src/META-INF/SAML.xml
```

This file is **not included** in the repository because it contains your tenant-specific SAML envelope. Create it at that path with the following structure, replacing the `[...]` placeholders (the code substitutes them at runtime):

```xml
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
            xmlns:a="http://www.w3.org/2005/08/addressing"
            xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
  <s:Header>
    <a:Action>...</a:Action>
    <a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo>
    <a:To>[endpoint]</a:To>
    <o:Security xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
      <o:UsernameToken>
        <o:Username>[username]</o:Username>
        <o:Password>[password]</o:Password>
      </o:UsernameToken>
    </o:Security>
  </s:Header>
  <s:Body>
    ...
  </s:Body>
</s:Envelope>
```

The `[username]`, `[password]`, and `[endpoint]` tokens are replaced with the values passed to the `SharePointClient` constructor.

> **Security note:** Never commit `SAML.xml` (or any file containing credentials) to version control. Add it to `.gitignore`.

---

## Troubleshooting

### `wsimport: command not found`
`wsimport` was removed from JDK 9+. Install JDK 8 and add its `bin` directory to the **front** of your `PATH`. See [wsimport note](#wsimport-note).

### `[ERROR] directory not found: target/generated-sources/wsimport`
Run `mvn initialize` once to create the directory, then retry `mvn test`. This should not occur in normal usage.

### `BUILD FAILURE` — compilation errors referencing `javax.xml.ws.*`
This means the JAX-WS RI dependency (`com.sun.xml.ws:rt`) was not resolved. Check your internet connection and run `mvn dependency:resolve` to download it from Maven Central.

### `java.io.FileNotFoundException` for `SAML.xml` at runtime
The SAML template file is missing. Create `src/META-INF/SAML.xml` as described in [Authentication: SAML Template](#authentication-saml-template).

### `Invalid CookieNed Token!!!` exception
SharePoint did not return `FedAuth` and `rtFa` cookies. This usually means:
- Wrong username or password.
- The SAML XML template is malformed or incomplete.
- Multi-factor authentication is enforced on the account (app passwords or a service account may be required).

### Tests fail on Windows with encoding errors
Ensure the terminal and the JVM use UTF-8. Add `-Dfile.encoding=UTF-8` to the Maven command:
```cmd
mvn test -Dfile.encoding=UTF-8
```
