function add_rels(rels, rId, f, type, relobj) {
	if(!relobj) relobj = {};
	if(!rels['!id']) rels['!id'] = {};
	relobj.Id = 'rId' + rId;
	relobj.Type = type;
	relobj.Target = f;
	if(rels['!id'][relobj.Id]) throw new Error("Cannot rewrite rId " + rId);
	rels['!id'][relobj.Id] = relobj;
	rels[('/' + relobj.Target).replace("//","/")] = relobj;
}

function write_zip(wb, opts) {
	if(wb && !wb.SSF) {
		wb.SSF = SSF.get_table();
	}
	if(wb && wb.SSF) {
		make_ssf(SSF); SSF.load_table(wb.SSF);
		opts.revssf = evert_num(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
	}
	opts.rels = {}; opts.wbrels = {};
	opts.Strings = []; opts.Strings.Count = 0; opts.Strings.Unique = 0;
	var wbext = opts.bookType == "xlsb" ? "bin" : "xml";
	var ct = { workbooks: [], sheets: [], calcchains: [], themes: [], styles: [],
		coreprops: [], extprops: [], custprops: [], strs:[], comments: [], vba: [],
		TODO:[], rels:[], xmlns: "", drawings: [] };
	fix_write_opts(opts = opts || {});
	var zip = new jszip();
	var f = "", rId = 0;

	opts.cellXfs = [];
	get_cell_style(opts.cellXfs, {}, {revssf:{"General":0}});

	f = "docProps/core.xml";
	zip.file(f, write_core_props(wb.Props, opts));
	ct.coreprops.push(f);
	add_rels(opts.rels, 2, f, RELS.CORE_PROPS);

	f = "docProps/app.xml";
	if(!wb.Props) wb.Props = {};
	wb.Props.SheetNames = wb.SheetNames;
	wb.Props.Worksheets = wb.SheetNames.length;
	zip.file(f, write_ext_props(wb.Props, opts));
	ct.extprops.push(f);
	add_rels(opts.rels, 3, f, RELS.EXT_PROPS);

	if(wb.Custprops !== wb.Props && keys(wb.Custprops||{}).length > 0) {
		f = "docProps/custom.xml";
		zip.file(f, write_cust_props(wb.Custprops, opts));
		ct.custprops.push(f);
		add_rels(opts.rels, 4, f, RELS.CUST_PROPS);
	}
    
    if(typeof wb.Images == 'object' && wb.Images instanceof Array) {
        for(var imgId = 0, len = wb.Images.length; imgId < len; ++imgId) {
            zip.file('xl/media/image' + (imgId + 1) + '.' + wb.Images[imgId].type, wb.Images[imgId].data, {base64: true});
        }
    }
    
    if(typeof wb.Drawings == 'object' && wb.Drawings instanceof Array) {
        for(var drawId = 0, len = wb.Drawings.length; drawId < len; ++drawId) {
            f = 'xl/drawings/drawing' + (drawId + 1) + '.xml';
            zip.file(f, write_drawing(wb.Drawings[drawId]));
            zip.file('xl/drawings/_rels/drawing' + (drawId + 1) + '.xml.rels', write_drawing_rels(wb, wb.Drawings[drawId]));
            ct.drawings.push(f);
        }
    }

	f = "xl/workbook." + wbext;
	zip.file(f, write_wb(wb, f, opts));
	ct.workbooks.push(f);
	add_rels(opts.rels, 1, f, RELS.WB);

	for(rId=1;rId <= wb.SheetNames.length; ++rId) {
		f = "xl/worksheets/sheet" + rId + "." + wbext;
		zip.file(f, write_ws(rId-1, f, opts, wb));
		ct.sheets.push(f);
		add_rels(opts.wbrels, rId, "worksheets/sheet" + rId + "." + wbext, RELS.WS);
        var ws = wb.Sheets[wb.SheetNames[rId-1]];
        var wsrels = {};
        if(ws['!drawing'] != undefined)
        {
            add_rels(wsrels, 1, '../drawings/drawing'+(ws['!drawing']+1)+'.xml', RELS.DRAWING);
        }
        if(Object.keys(wsrels).length > 0)
        {
            zip.file('xl/worksheets/_rels/sheet' + rId + '.xml.rels', write_rels(wsrels));
        }
	}

	if(opts.Strings != null && opts.Strings.length > 0) {
		f = "xl/sharedStrings." + wbext;
		zip.file(f, write_sst(opts.Strings, f, opts));
		ct.strs.push(f);
		add_rels(opts.wbrels, ++rId, "sharedStrings." + wbext, RELS.SST);
	}

	/* TODO: something more intelligent with themes */

	f = "xl/theme/theme1.xml";
	zip.file(f, write_theme(opts));
	ct.themes.push(f);
	add_rels(opts.wbrels, ++rId, "theme/theme1.xml", RELS.THEME);

	/* TODO: something more intelligent with styles */

	f = "xl/styles." + wbext;
	zip.file(f, write_sty(wb, f, opts));
	ct.styles.push(f);
	add_rels(opts.wbrels, ++rId, "styles." + wbext, RELS.STY);

	zip.file("[Content_Types].xml", write_ct(ct, opts));
	zip.file('_rels/.rels', write_rels(opts.rels));
	zip.file('xl/_rels/workbook.' + wbext + '.rels', write_rels(opts.wbrels));
	return zip;
}
