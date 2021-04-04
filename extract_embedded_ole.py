#!/usr/bin/env python3

import olefile
import os

def extract_embedded_ole(fname):
    ole = olefile.OleFileIO(fname)
    stream_types = ['Workbook', 'WordDocument', 'Package', 'WordDocument',
                    'VisioDocument', 'PowerPoint Document',
                    "Book", "CONTENTS"
                    ]
    i = 0
    for stream in ole.listdir():
        for s in stream:
            if isinstance(stream, list) and len(stream) > 1:
                i += 1
            check_data_stream = False
            if ole.get_type(stream) == 2:
                check_data_stream = True
            if s in stream_types and check_data_stream:
                ole_stream = ole.openstream(stream)
                ole_props = ole.getproperties(['\x05SummaryInformation'])
                out_dir = fname + ".embeddings/" + "/".join(stream[:-1])
                try:
                    os.makedirs(out_dir)
                except OSError as e:
                    print(e)
                # Write out Streams
                out_name = out_dir + "/" + os.path.split(fname)[1]
                embed_name = "-emb-" + s + "-" + str(i) + ".ole"
                out_name += embed_name
                print(out_name)
                out_name = out_name + embed_name
                out_file = open(out_name, 'w+b')
                out_file.write(ole_stream.read())
    out_file.close()

extract_embedded_ole('winter_extended.doc')
