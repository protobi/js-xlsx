function write_zip_type(wb, opts) {
	var o = opts||{};
  style_builder  = new StyleBuilder(opts);

  var z = write_zip(wb, o);
    var oopts = {};
    if(o.compression) oopts.compression = 'DEFLATE';
	switch(o.type) {
		case "base64": oopts.type = "base64"; break;
		case "binary": oopts.type = "string"; break;
		case "buffer":
		case "file": oopts.type = "nodebuffer"; break;
		default: throw new Error("Unrecognized type " + o.type);
	}
    var out = z.generate(oopts);
    if(o.type === "file") return _fs.writeFileSync(o.file, out);
    return out;
}

function writeSync(wb, opts) {
	var o = opts||{};
	switch(o.bookType) {
		case 'xml': return write_xlml(wb, o);
		default: return write_zip_type(wb, o);
	}
}

function writeFileSync(wb, filename, opts) {
	var o = opts||{}; o.type = 'file';

	o.file = filename;
	switch(o.file.substr(-5).toLowerCase()) {
		case '.xlsx': o.bookType = 'xlsx'; break;
		case '.xlsm': o.bookType = 'xlsm'; break;
		case '.xlsb': o.bookType = 'xlsb'; break;
	default: switch(o.file.substr(-4).toLowerCase()) {
		case '.xls': o.bookType = 'xls'; break;
		case '.xml': o.bookType = 'xml'; break;
	}}
	return writeSync(wb, o);
}

