RELS.DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';

var write_drawing = (function()
{
    var DOC_START = '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
    var DOC_END = '</xdr:wsDr>';
    var ANCH_START = '<xdr:twoCellAnchor editAs="oneCell">';
    var ANCH_END = '</xdr:twoCellAnchor>';
    
    var ONE = '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="';
    var TWO = '" name=""/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId';
    var THREE = '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>';
    
    function write_coords(name, data)
    {
        var o = ['<xdr:' + name + '>'];
        o[o.length] = '<xdr:col>' + data.col + '</xdr:col>';
        o[o.length] = '<xdr:colOff>' + data.colOff + '</xdr:colOff>';
        o[o.length] = '<xdr:row>' + data.row + '</xdr:row>';
        o[o.length] = '<xdr:rowOff>' + data.rowOff + '</xdr:rowOff>';
        o[o.length] = '</xdr:' + name + '>';
        return o.join("");
    }
    
    return function(data)
    {
        var o = [XML_HEADER];
        o[o.length] = DOC_START;
        data.forEach(function(e, i)
        {
            o[o.length] = ANCH_START;
            o[o.length] = write_coords('from', e.from);
            o[o.length] = write_coords('to', e.to);
            o[o.length] = ONE + (i + 1) + TWO + (i + 1) + THREE;
            o[o.length] = ANCH_END;
        });
        o[o.length] = DOC_END;
        return o.join("");
    }
})();

function write_drawing_rels(wb, data)
{
    var o = [XML_HEADER];
    o[o.length] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    data.forEach(function(e, i)
    {
        var imgpath = '../media/image' + (e.image + 1) + '.' + wb.Images[e.image].type;
        o[o.length] = '<Relationship Id="rId'+(i+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + imgpath + '"/>';
    });
    o[o.length] = '</Relationships>';
    return o.join("");
}