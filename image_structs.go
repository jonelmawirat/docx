package docx

import "encoding/xml"


type DocProperties struct {
    XMLName xml.Name `xml:"wp:docPr"`
    ID      uint     `xml:"id,attr"`
    Name    string   `xml:"name,attr"`
    Descr   string   `xml:"descr,attr,omitempty"`
}


type graphicFrameLocks struct {
    XMLName        xml.Name `xml:"a:graphicFrameLocks"`
    XmlnsA         string   `xml:"xmlns:a,attr,omitempty"`
    NoChangeAspect uint     `xml:"noChangeAspect,attr"`
}


type CnvGraphicFrameProperties struct {
    XMLName      xml.Name          `xml:"wp:cNvGraphicFramePr"`
    GraphicFrame graphicFrameLocks `xml:"a:graphicFrameLocks"`
}


type point2D struct {
    XMLName xml.Name `xml:"a:off"`
    X       int64    `xml:"x,attr"`
    Y       int64    `xml:"y,attr"`
}


type extents struct {
    XMLName xml.Name `xml:"a:ext"`
    Cx      int64    `xml:"cx,attr"`
    Cy      int64    `xml:"cy,attr"`
}


type transform2D struct {
    XMLName xml.Name `xml:"a:xfrm"`
    Offset  point2D  `xml:"a:off"`
    Extents extents  `xml:"a:ext"`
}


type presetGeometry struct {
    XMLName xml.Name `xml:"a:prstGeom"`
    Prst    string   `xml:"prst,attr"`
    AVList  *struct {
        XMLName xml.Name `xml:"a:avLst"`
    } `xml:"a:avLst"`
}


type blip struct {
    XMLName xml.Name `xml:"a:blip"`
    Embed   string   `xml:"r:embed,attr"`
    Cstate  string   `xml:"cstate,attr,omitempty"`
}


type srcRect struct {
    XMLName xml.Name `xml:"a:srcRect"`
}


type stretch struct {
    XMLName       xml.Name `xml:"a:stretch"`
    FillRectangle *struct {
        XMLName xml.Name `xml:"a:fillRect"`
    } `xml:"a:fillRect"`
}


type blipFill struct {
    XMLName xml.Name `xml:"pic:blipFill"`
    Blip    blip     `xml:"a:blip"`
    SrcRect *srcRect `xml:"a:srcRect,omitempty"`
    Stretch stretch  `xml:"a:stretch"`
}


type noFill struct {
    XMLName xml.Name `xml:"a:noFill"`
}


type headEnd struct {
    XMLName xml.Name `xml:"a:headEnd"`

}

type tailEnd struct {
    XMLName xml.Name `xml:"a:tailEnd"`

}


type ln struct {
    XMLName xml.Name `xml:"a:ln"`
    NoFill  *noFill  `xml:"a:noFill,omitempty"`
    Miter   *struct {
        XMLName xml.Name `xml:"a:miter"`
        Lim     string   `xml:"lim,attr"`
    } `xml:"a:miter,omitempty"`
    HeadEnd *headEnd `xml:"a:headEnd,omitempty"`
    TailEnd *tailEnd `xml:"a:tailEnd,omitempty"`
}


type shapeProperties struct {
    XMLName  xml.Name       `xml:"pic:spPr"`
    BwMode   string         `xml:"bwMode,attr,omitempty"`
    Xfrm     transform2D    `xml:"a:xfrm"`
    PrstGeom presetGeometry `xml:"a:prstGeom"`
    NoFill   *noFill        `xml:"a:noFill,omitempty"`
    Ln       *ln            `xml:"a:ln,omitempty"`
}


type picLocks struct {
    XMLName            xml.Name `xml:"a:picLocks"`
    NoChangeAspect     uint     `xml:"noChangeAspect,attr,omitempty"`
    NoChangeArrowheads uint     `xml:"noChangeArrowheads,attr,omitempty"`
}


type cNvPicPr struct {
    XMLName  xml.Name `xml:"pic:cNvPicPr"`
    PicLocks picLocks `xml:"a:picLocks"`
}


type nonVisualPicProperties struct {
    XMLName xml.Name `xml:"pic:nvPicPr"`
    CNvPr   struct {
        XMLName xml.Name `xml:"pic:cNvPr"`
        ID      uint     `xml:"id,attr"`
        Name    string   `xml:"name,attr"`
        Descr   string   `xml:"descr,attr,omitempty"`
    } `xml:"pic:cNvPr"`
    CNvPicPr cNvPicPr `xml:"pic:cNvPicPr"`
}


type pic struct {
    XMLName  xml.Name               `xml:"pic:pic"`
    XmlnsPic string                 `xml:"xmlns:pic,attr,omitempty"`
    NvPicPr  nonVisualPicProperties `xml:"pic:nvPicPr"`
    BlipFill blipFill               `xml:"pic:blipFill"`
    SpPr     shapeProperties        `xml:"pic:spPr"`
}


type graphicData struct {
    XMLName xml.Name `xml:"a:graphicData"`
    URI     string   `xml:"uri,attr"`
    Pic     pic      `xml:"pic:pic"`
}


type graphic struct {
    XMLName     xml.Name    `xml:"a:graphic"`
    XmlnsA      string      `xml:"xmlns:a,attr,omitempty"`
    GraphicData graphicData `xml:"a:graphicData"`
}


type extent struct {
    XMLName xml.Name `xml:"wp:extent"`
    Cx      int64    `xml:"cx,attr"`
    Cy      int64    `xml:"cy,attr"`
}


type effectExtent struct {
    XMLName xml.Name `xml:"wp:effectExtent"`
    L       int64    `xml:"l,attr"`
    T       int64    `xml:"t,attr"`
    R       int64    `xml:"r,attr"`
    B       int64    `xml:"b,attr"`
}


type Inline struct {
    XMLName xml.Name `xml:"wp:inline"`
    DistT   uint     `xml:"distT,attr"`
    DistB   uint     `xml:"distB,attr"`
    DistL   uint     `xml:"distL,attr"`
    DistR   uint     `xml:"distR,attr"`


    Extent            extent                  `xml:"wp:extent"`
    EffectExtent      effectExtent            `xml:"wp:effectExtent"`
    DocPr             DocProperties           `xml:"wp:docPr"`
    CNvGraphicFramePr CnvGraphicFrameProperties `xml:"wp:cNvGraphicFramePr"`


    Graphic graphic `xml:"a:graphic"`
}


type Drawing struct {
    XMLName xml.Name `xml:"w:drawing"`
    Inline  Inline   `xml:"wp:inline"`
}
