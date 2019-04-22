package com.zc.officereader.wordreader;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class CellInfo {
    Integer XPosition;
    Integer YPosition;
    String content;
    Boolean isTitle;
}
