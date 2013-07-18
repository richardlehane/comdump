Take a quick look at MS compound objects.

Tool based on github.com/richardlehane/mscfb package.
It creates files for each of the directory entries in an compound object
and writes them to a comobjects directory. Extracts JPGs from Thumbs.db
files if you add a -thumbs switch.

Examples:

    ./comdump -in test.doc
    ./comdump -in Thumbs.db -thumbs
 