Take a quick look inside MS compound file binary file format (OLE2/COM) files.

Tool based on github.com/richardlehane/mscfb package.
It creates files for each of the directory entries in an compound object
and writes them to a comobjects directory.
Extracts metadata (CLSID and date information) with the -meta switch.
Extracts JPGs and Catalog metadata from Thumbs.db files with the -thumbs switch.

Examples:

    ./comdump test.doc
    ./comdump -meta test.doc
    ./comdump -thumbs Thumbs.db 
 
 Install with `go get` and compile. Or download one of the binary releases.
