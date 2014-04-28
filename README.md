SynProject
==========

*Synopse SynProject* is an open source application for code source versioning and automated documentation of Delphi projects.

Licensed under a GPL license.

[Take a look at the SynProject official web site](http://synopse.info/fossil/wiki?name=SynProject)



Features
========

Its main features are:

1. Local source code versioning;
2. Automated documentation.


Source code versioning
----------------------

  * handle multiple projects or libraries with the same program;
  * allow source code versioning with detailed commits;
  * can access to a PVCS tracker (more trackers are coming) and link the commits to the tickets;
  * allow automated source code backup (without any commit to document: it's like a daily snapshot of your files);
  * backups can be local (on your hard drive) or remote (on a distant drive);
  * you can make a visual "diff" and compare source code versions side by side in the graphical user interface;
  * you can see pictures (png jpeg bitmap icon) within the main user interface;
  * diff storage between version is very optimized, and use little disk space;
  * storage is based on .zip files and plain text files, so it's easy to work with.

Automated Documentation
-----------------------

  * follow a typical Design Inputs -> Risk Assessment -> Software Architecture -> Detailed Design -> Tests protocols -> Traceability matrix -> Release Notes workflow;
  * initial (marketing-level) Design Input can be refined into more precise Software Requirement Specifications;
  * Design Inputs can evolve during the project life: all documentation stay synchronized and you will have to maintain the DI and their description only at one place;
  * therefore, the process is meant to be compliant with the most precise documentation protocols (like IEC 62304);
only one text file, formated like a wiki, contains the whole documentation;
  * it's very easy to add pictures, or formated source code (pascal, C, C++, C#, plain text) into the documentation;
  * word files (and then pdf) are created from this content, with full table of contents, picture or source code reference tables, unified page layout, customizable templates;
  * it's easy to add tables to your document, or link to other part or external resources * you can even put pure RTF content into your documentation;
  * pictures are centralized and captioned, people involved in the documentation are maintained once for the whole documentation;
  * document version numbering and cross-referencing is handled easily;
  * for pascal projects, the source code is parsed and all interface architecture is generated from the source;
it's easy to browse classes, variables and functions from the documentation, and add reference to them to your document;
the source code description can also be located in an external .sae file, therefore your original source code tree won't necessary be changed by the adding of comments;
  * all references are cross-linked: Software Architecture Document is created from the source code, and also is able to highlight the classes or methods - involved in implementing every Design Input, from the Software Design Document;
  * integrated GraphViz component, in order to create easily diagrams from plain text embedded into your documentation;   * integrated fully featured text editor, with word wrapping, wiki-syntax buttons and keyboard shortcuts, and spell checking;
  * a step by step Wizard is available to create a new project, from supplied template files;
another Wizard is already available, to browse your documentation workflow, and check its consistency or set up its parameters;
  * the documentation is integrated to the Versioning system above.

Of course, it's a full Open Source - GPL licensed - project. Source code (for Delphi 6/7) is available and maintained, since it is used by our internal projects (including our little [mORMot](http://mormot.net)).