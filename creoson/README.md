# CREOSON

CREOSON is an OpenSource initiative by [Simplified Logic, Inc.](http://www.simplifiedlogic.com) to promote automation of PTC's CREO Parametric.

CREOSON uses simple JSON Requests to send commands/functions to CREO, JSON Responses are used to communicate the results of your requests.



# Pre-Packaged Download(s)!
If you are not a "developer"... and don't want to build CREOSON from "scratch" ... a Pre-Packaged Distribution is available for 32-bit or 64-bit CREO.  The Pre-Packaged Distributions have everything (including the correct version of Java) already baked-in and ready to rock!

You can literally download, unzip and run CREOSON Micro-Server immediately!

 Pre-Packaged Distributions are located here:
 - --> [Download CRESOSON Releases](https://github.com/SimplifiedLogic/creoson/releases) <--



### Minimum Requirements

CREOSON Requires the following to work:

* CREO Parametric - Version 2.0 or Later... 3.0/4.0 have most coverage/stability.
* JLINK - This is a FREE Interface with CREO Parametric, but must be installed!
* Permissions on your machine - don't forget to check your firewall!

* A "little" programming/scripting knowledge does not hurt.



### Installation?

There is nothing to "install", the Pre-Packaged Distribution contains everything you need, but here are some basic steps to follow:

1. Download your Pre-Packaged Distribution  (Distribution w/ GUI Utility is Recommended -> CreosonServerWithSetup-*.zip)
2. Move the ZIP file to a directory where you want CREOSON to run from
3. UNZIP the Pre-Packaged Distribution
4. Run the CreosonSetup.exe
5. Select your CREO Installation Directory
6. Select a PORT Number for the CREOSON Server to use
7. Press "Start CREOSON"
- Creoson is now running!
8. Press "Open Documentation"
- We suggest you check out the "Playground" to play with a few commands, then jump over to the "Functions" section and dig into the examples.


# Need Automation Help?
----

Visit the --> [PTC Creo Parametric Community](https://community.ptc.com/t5/Creo-Parametric/ct-p/creoparametric) <-- to learn more ...

If you need professional help with your automation, send us an email with your goals and requirements to --> creoson@simplifiedlogic.com <--

# Technical Details

----

### Why JSON?!

JSON was chosen because it is super generic and easy to work with.

Typically, CREO Parametric automation options are limited to a specific set of core languages like C/C++/C#, Java and Visual Basic.  JSON enables you to use ANY programming/scripting language to manipulate CREO Parametric via CREOSON.  So you can use scripting languages like Ruby, Python, PHP or even Web Technologies like AJAX to perform operations in CREO quite easily - and faster!


### The Geeky CREOSON Details

CREOSON uses a Micro-Server to "listen" for JSON requests and then handles/translates them into JLINK calls via our long-used "jShell" interface to JLINK.  jShell is a highly optimized business/logic-interface to JLINK and CREOSON is the JSON interface to jShell.  If it is confusing... well.. you are reading the "geeky" section... DON'T worry - most users will only need to use Pre-Packaged Distribution and learn how to use JSON Transactions to do amazing things quickly.



# SOURCE CODE
----
Sub-project folders -- these contain the project source and Ant build script for each Creoson jar file:

* creofuncs
* creoson-core
* creoson-intf
* creoson-json
* creoson-json-const
* creoson-server -- also contains the website files

### Files in the root directory

* BUILD_README.txt -- short doc explaining how to build the project from scratch
* build_vars.properties -- properties file used for builds which contain settings that the users are going to have to set manually -- the Creo location and the location of the CreosonSetup application.
* The various root-level build files -- build-javadoc.xml, build-zip.xml, build-all.xml
* RELEASE_NOTES.txt -- release notes
* creoson_source_doc.pdf -- high-level documentation about how the source code is organised



# Contribute!
----
Want to contribute? Great!

Sign up with GitHub, pull the source and start to help!



# License
----

MIT

