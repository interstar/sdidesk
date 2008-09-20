
Welcome to SdiDesk 0.2.2 sourcecode (scroll down for diffs with 0.2.0)

The code is, I hope, much improved from version 0.1


Quick start
-----------

It's VB. The main project file is called DesktopWiki.vbp


What's new in 0.2?
------------------

1) The Architecture of Exporting =======================================================

The most important change is that I stripped out any pretence of exporting pages from the SdiDesk program.

SdiDesk is a tool to organize you're thoughts and make notes. But in an ideal world, you'll be able to export in a variety of formats, allowing you to use it to prepare documents. 

That's still the aim, but it was getting fiendishly complex. 

So, I've decided to split the exporters out from SdiDesk into separate programs that live in the "exporters" sub-directory.

That has several virtues. 

- I can get this release of SdiDesk out without worrying any more about exporting

- The main SdiDesk code has lost the ugliest most confusing stuff

- You can add new exporters to your installation at any time, simply by dropping the appropriate program into the exporters directory

- Anyone else can write and distribute exporters without needing me to compile them into the main code-base release.

- Instant "Architecture of Participation"

How does it work?

SdiDesk now uses the Shell command to call exporters. 

It sends command arguments like this :

For "instant" exports (when you hit the "export" button on a page)

 exporterName.exe -page 'PageName' -psi 'C:\sdidesk\'
 
The "page" argument is the page to be exported. The "psi" argument stands for PageStore Identifier. At the moment this just means the directory where SdiDesk is running. (So the exporter can find the PageStore)

If the exporter needs other information, it must get it from the user.

When running an "export", ie. exporting a number of pages, SdiDesk calls with 

 exporterName.exe -name 'ExportName' -param 'ParamPage' -psi 'C:\...'


Here, name is the name of the export, param is the name of a page where exporter can look for further parameters, and psi is, again, the page-store identifier.

Any further information that might need to be passed to the exporter, for example : which PageCrawler to use to pick up the pages to be exported, which page the crawler will start it's crawl on, which page contains a template etc. is expected to be on the parameter page in the form of Page variables. 

eg. if ParamPage contains the line : 

start #= HomePage

this will create a Page variable on ParamPage, who's name is "start" and who's value is "HomePage". An exporter looking on ParamPage will be able to find this info.

As of now, I'm working on getting the exportation of flat web-sites done in this new style. The Work in Progress is on the program called HtmlExporter. It's not working or finished, but I left the source code in for you to see what's going on. Exporting is now going to be heavily following the two interfaces : "Export" and "Exporter".

DummyExporter is a working program which simply picks up the command line arguments and can also get the paramaters from a param page. It shows you how it's done.


2) Interface Frenzy =====================================================================

Nearly anything of any importance is behind an interface these days. 

SdiDesk is still (almost cleanly) separated into Model, View and Controller. There are a few 
places it isn't, but I'm working on.

But the complexity was spiraling out of control. Hence I've been slotting in new interfaces 
all over, to carve it into managable parts. 

Most of the work was being done by the object I used to call ModelLevel. (It was the model of the MVC split.) 

But it was becoming a huge "ball of mud". In particular, almost everything needed to refer back to it to get pages from the PageStore, or to find configuration info. or for dependency injection. That meant it was nearly impossible to break bits of SdiDesk out for separate testing. Everything was interdependent.

What I did, was divide this work of the model between a number of interfaces :

- WikiAnnotatedDataStore 

  which basically wraps PageStore and some other stuff

- SingleUserState 

  which handles all the things relevant to the user's session, eg. navigation history, whether automatic backlinks are on, etc

- SystemConfigurations

  configurations of the system in general. The InterMap, the colours for link-types 

- CrawlerSubsystem

  Anything which engages PageCrawlers

- ExportSubsystem

  All the functions needed to trigger an export

- ControllableModel

  This is the interface presented to the ControlLevel. The ControlLevel still does the work of firing off other behaviours from the ModelLevel. Most of the time it does it via this interface.


- LocalFileSystem

When you need to see the local file-system.

Eventually, these interfaces could be wrapped around separate classes. At the moment, the ModelImplementation handles them all.

A lot of other things are now interfaces, here's a list with the concrete implementation.

- Page, MemoryResidentPage
- PageStore, FileSystemPageStore
- PageCooker, NativePageCooker


3) Networks are Pages ==================================================================

In the old version, a network was a component of Page. This was, frankly, bad, in that every program which touched pages (almost any test program), also had to know about networks. Now a network is more like a "decorator". It also implements the Page interface and contains a page. When a page isNetwork, the model's getPage creates a Network as a wrapper and returns that, with the page inside. All Page functions are delegated to the inner page.

Later on, I *might* do something similar with Tables, though it seems less urgent.

4) Table editing uses a grid control ====================================================

Tables now have their own special editor, as networks do. And this uses the FlexGrid control. This is pretty primitive at the moment (doesn't let you add new columns or rows) but should grow.

5) Network editing has introduced new pop-up windows ====================================

There's a good and bad side to this. I *hate* pop-up windows. A major part of the philosophy 
of SdiDesk is to avoid them. But in this case, they seemed the best option. Don't like 'em though. Suggestions for alternatives are welcome.

6) New Features =========================================================================

- wordCount
- footnotes
- calendars
- you can add buttons to pages

Cute for users, nothing very dramatic from the code perspective.
 

OK, I'll add more documentation on the wiki. Hopefully there'll be some more incremental 
releases to fix bugs, tidy up the loose ends and get a couple of exporters out over the next few weeks.

-- PhilJones (May, 2005)


------------------------------
Diffs 0.2.1 -> 0.2.2

- fixed bug with links to external sites

- fixed bug with buttons on pages


Diffs 0.2.0 -> 0.2.1

- fixed bug in wordCount (isAlpha doesn't die if it receives an empty argument)

- non-existent pages default to "edit" mode

- list elements can be highlighted if final char is *

- explicit references to NativePageCooker replaced by PageCooker and some refactoring