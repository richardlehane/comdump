// Copyright 2013 Richard Lehane. All rights reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

// Take a quick look inside MS compound file binary file format (OLE2/COM) files.

// Tool based on github.com/richardlehane/mscfb package.
// It creates files for each of the directory entries in an compound object
// and writes them to a comobjects directory.
// Extracts JPGs from Thumbs.db files if you add a -thumbs switch.
//
// Examples:
//    ./comdump test.doc
//    ./comdump -thumbs Thumbs.db
package main

import (
	"flag"
	"github.com/richardlehane/mscfb"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
	"unicode"
)

var thumbs = flag.Bool("thumbs", false, "treat input Comound object as Thumbs.db file")

func clean(str string) string {
	buf := make([]rune, 0, len(str))
	for _, r := range str {
		if unicode.IsPrint(r) {
			buf = append(buf, r)
		}
	}
	return string(buf)
}

func process(in string, thumbs bool) error {
	thumbsBuf := make([]byte, 24)

	file, err := os.Open(in)
	if err != nil {
		return err
	}
	defer file.Close()

	doc, err := mscfb.NewReader(file)
	if err != nil {
		return err
	}
	dir, base := filepath.Split(in)
	base = strings.Join(strings.Split(base, "."), "_")
	base += "_comobjects"
	path := filepath.Join(dir, base)
	err = os.Mkdir(path, os.ModePerm)
	if err != nil {
		return err
	}

	for {
		entry, err := doc.Next()
		if err == io.EOF {
			break
		}
		if entry.Dir {
			continue
		}
		if err != nil {
			return err
		}
		name, names := doc.Name()
		if len(names) > 1 {
			names = append(names[1:], name)
		} else {
			names = []string{name}
		}
		for _, v := range names {
			clean(v)
		}
		if thumbs {
			names[len(names)-1] += ".jpg"
			_, err = doc.Read(thumbsBuf)
			if err != nil {
				return err
			}
		}
		outFile, err := os.Create(filepath.Join(path, names...))
		if err == nil {
			_, err = io.Copy(outFile, doc)
		}
		outFile.Close()
		if err != nil {
			return err
		}
	}
	return nil
}

func main() {
	flag.Parse()
	ins := flag.Args()
	if len(ins) < 1 {
		log.Fatalln("Missing required argument: path_to_compound_object")
	}
	for _, in := range ins {
		err := process(in, *thumbs)
		if err != nil {
			log.Fatalln(err)
		}
	}
}
