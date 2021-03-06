package com

import java.util.stream.Collectors

class DefaultCitiDataReader implements CitiDataReader {

    List<CitiData> read(File dataFile) {
        def text = dataFile.text
        text.split('(\r\n|\n)').toList().stream().map {
            def fields = it.split(',')
            new CitiData(date: fields[0], id: fields[1])
        }.collect(Collectors.toList())
    }
}
