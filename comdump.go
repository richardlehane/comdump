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
	"encoding/binary"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf16"

	"github.com/richardlehane/mscfb"
	"github.com/richardlehane/msoleps/types"
)

var debug = flag.Bool("debug", false, "print debugging info")

var thumbs = flag.Bool("thumbs", false, "treat input Compound object as Thumbs.db file")

var meta = flag.Bool("meta", false, "output Compound object metadata")

type CatalogHeader struct {
	Magic      uint16
	Version    uint16
	NumEntries uint32
	DimensionX uint32
	DimensionY uint32
}

type CatalogEntry struct {
	Size   uint32
	Number uint32
	Date   types.FileTime
}

func process(in string, thumbs bool) error {
	thumbsSz := make([]byte, 4)

	file, err := os.Open(in)
	if err != nil {
		return err
	}
	defer file.Close()

	doc, err := mscfb.New(file)
	if doc == nil {
		return err
	}
	if *debug {
		d := doc.Debug()
		fmt.Println("DEBUGGING")
		for k, v := range d {
			fmt.Printf("%s: %v\n", k, v)
		}
		return nil
	}
	dir, base := filepath.Split(in)
	base = strings.Join(strings.Split(base, "."), "_")
	base += "_comobjects"
	path := filepath.Join(dir, base)
	if *meta {
		if err != nil {
			fmt.Println("Errors: ", err.Error())
		}
		fmt.Println("Root Object")
		fmt.Println("  CLSID:     ", doc.ID())
		fmt.Println("  Created:   ", doc.Created())
		fmt.Println("  Modified:  ", doc.Modified())
		fmt.Println()
	} else {
		err = os.Mkdir(path, os.ModePerm)
		if err != nil {
			return err
		}
	}
	for {
		entry, err := doc.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		paths := []string{path}
		entry.Path = append(entry.Path, entry.Name)
		paths = append(paths, entry.Path...)
		if *meta {
			if !entry.FileInfo().IsDir() {
				fmt.Println("Stream Object")
				fmt.Println("  Name :     ", entry.Name)
				fmt.Println("  Initial:   ", entry.Initial)
				fmt.Println("  Path:      ", strings.Join(entry.Path, "/"))
				fmt.Printf("  Size:       %d", entry.Size)
			} else {
				fmt.Println("Storage Object")
				fmt.Println("  Name (raw):", entry.Name)
				fmt.Println("  Path:      ", strings.Join(entry.Path, "/"))
				fmt.Println("  CLSID:     ", entry.ID())
				fmt.Println("  Created:   ", entry.Created())
				fmt.Println("  Modified:  ", entry.Modified())
			}
			fmt.Println()
			continue
		}
		if entry.FileInfo().IsDir() {
			err = os.Mkdir(filepath.Join(paths...), os.ModePerm)
			if err != nil {
				return err
			}
			continue
		}
		if thumbs && entry.Name != "Catalog" {
			paths[len(paths)-1] += ".jpg"
			_, err = doc.Read(thumbsSz)
			if err != nil {
				return err
			}
			sz := binary.LittleEndian.Uint32(thumbsSz)
			cut := make([]byte, int(sz)-4)
			_, err = doc.Read(cut)
			if err != nil {
				return err
			}
		}
		if thumbs && entry.Name == "Catalog" {
			hdr := new(CatalogHeader)
			err = binary.Read(doc, binary.LittleEndian, hdr)
			if err != nil {
				return err
			}
			fmt.Println("Thumbs Database")
			fmt.Println("  Version:    ", hdr.Version)
			fmt.Println("  Thumbnails: ", hdr.NumEntries)
			fmt.Println("  DimensionX: ", hdr.DimensionX)
			fmt.Println("  DimensionY: ", hdr.DimensionY)
			for i := 0; i < int(hdr.NumEntries); i++ {
				thumb := new(CatalogEntry)
				err = binary.Read(doc, binary.LittleEndian, thumb)
				if err != nil {
					return err
				}
				name := make([]uint16, (int(thumb.Size)-20)/2)
				err = binary.Read(doc, binary.LittleEndian, name)
				if err != nil {
					return err
				}
				fmt.Println("  Thumbnail ", thumb.Number)
				fmt.Println("    Name: ", string(utf16.Decode(name)))
				fmt.Println("    Date: ", thumb.Date)
				pad := make([]byte, 4)
				_, err = doc.Read(pad)
				if err != nil {
					return err
				}
			}
		}
		outFile, err := os.Create(filepath.Join(paths...))
		if err != nil {
			return err
		}
		_, err = io.Copy(outFile, doc)
		if err != nil {
			return err
		}
		outFile.Close()
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
