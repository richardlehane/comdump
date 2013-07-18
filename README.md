Take a quick look inside MS compound file binary file format (OLE2/COM) files.

Tool based on github.com/richardlehane/mscfb package.
It creates files for each of the directory entries in an compound object
and writes them to a comobjects directory.
Extracts JPGs from Thumbs.db files if you add a -thumbs switch.

Examples:

    ./comdump test.doc
    ./comdump -thumbs Thumbs.db 
 