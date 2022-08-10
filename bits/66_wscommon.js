var strs = {}; // shared strings
var _ssfopts = {}; // spreadsheet formatting options

RELS.WS = [
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	"http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
];

/*global Map */
var browser_has_Map = typeof Map !== 'undefined';

function get_sst_id(sst/*:SST*/, str/*:string*/, rev)/*:number*/ {
	var i = 0, len = sst.length;
	if(rev) {
		if(browser_has_Map ? rev.has(str) : Object.prototype.hasOwnProperty.call(rev, str)) {
			var revarr = browser_has_Map ? rev.get(str) : rev[str];
			for(; i < revarr.length; ++i) {
				if(sst[revarr[i]].t === str) { sst.Count ++; return revarr[i]; }
			}
		}
	} else for(; i < len; ++i) {
		if(sst[i].t === str) { sst.Count ++; return i; }
	}
	sst[len] = ({t:str}/*:any*/); sst.Count ++; sst.Unique ++;
	if(rev) {
		if(browser_has_Map) {
			if(!rev.has(str)) rev.set(str, []);
			rev.get(str).push(len);
		} else {
			if(!Object.prototype.hasOwnProperty.call(rev, str)) rev[str] = [];
			rev[str].push(len);
		}
	}
	return len;
}

function col_obj_w(C/*:number*/, col) {
	var p = ({min:C+1,max:C+1}/*:any*/);
	/* wch (chars), wpx (pixels) */
	var wch = -1;
	if(col.MDW) MDW = col.MDW;
	if(col.width != null) p.customWidth = 1;
	else if(col.wpx != null) wch = px2char(col.wpx);
	else if(col.wch != null) wch = col.wch;
	if(wch > -1) { p.width = char2width(wch); p.customWidth = 1; }
	else if(col.width != null) p.width = col.width;
	if(col.hidden) p.hidden = true;
	return p;
}

function default_margins(margins/*:Margins*/, mode/*:?string*/) {
	if(!margins) return;
	var defs = [0.7, 0.7, 0.75, 0.75, 0.3, 0.3];
	if(mode == 'xlml') defs = [1, 1, 1, 1, 0.5, 0.5];
	if(margins.left   == null) margins.left   = defs[0];
	if(margins.right  == null) margins.right  = defs[1];
	if(margins.top    == null) margins.top    = defs[2];
	if(margins.bottom == null) margins.bottom = defs[3];
	if(margins.header == null) margins.header = defs[4];
	if(margins.footer == null) margins.footer = defs[5];
}

function get_cell_style(styles/*:Array<any>*/, cell/*:Cell*/, opts) {
	if (typeof style_builder != 'undefined') {
		if (/^\d+$/.exec(cell.s)) { return cell.s }  // if its already an integer index, let it be
		if (cell.s && (cell.s == +cell.s)) { return cell.s }  // if its already an integer index, let it be
		var s = cell.s || {};
		if (cell.z) s.numFmt = cell.z;
		return style_builder.addStyle(s);
	} else {
		var z = opts.revssf[cell.z != null ? cell.z : "General"];
		var i = 0x3c, len = styles.length;
		if (z == null && opts.ssf) {
			for (; i < 0x188; ++i) if (opts.ssf[i] == null) {
				SSF.load(cell.z, i);
				// $FlowIgnore
				opts.ssf[i] = cell.z;
				opts.revssf[cell.z] = z = i;
				break;
			}
		}
		for (i = 0; i != len; ++i) if (styles[i].numFmtId === z) return i;
		styles[len] = {
			numFmtId: z,
			fontId: 0,
			fillId: 0,
			borderId: 0,
			xfId: 0,
			applyNumberFormat: 1
		};
		return len;
	}
}

function safe_format(p/*:Cell*/, cf_copy/*:CellXf*/, opts, themes, styles) {
	try {
		if (opts.cellNF) p.z = SSF._table[cf_copy.numFmtId];
	} catch (e) { if (opts.WTF) throw e; }
	// if (p.t === 'z') return;
	if (p.t === 'd' && typeof p.v === 'string') p.v = parseDate(p.v);
	if ((!opts || opts.cellText !== false) && p.v) try {
		if (SSF._table[cf_copy.numFmtId] == null) SSF.load(SSFImplicit[cf_copy.numFmtId] || "General", fmtid);
		if (p.t === 'e') p.w = p.w || BErr[p.v];
		else if (cf_copy.numFmtId === 0) {
			if (p.t === 'n') {
				if ((p.v | 0) === p.v) p.w = SSF._general_int(p.v);
				else p.w = SSF._general_num(p.v);
			}
			else if (p.t === 'd') {
				var dd = datenum(p.v);
				if ((dd | 0) === dd) p.w = SSF._general_int(dd);
				else p.w = SSF._general_num(dd);
			}
			else if (p.v === undefined) return "";
			else p.w = SSF._general(p.v, _ssfopts);
		}
		else if (p.t === 'd') p.w = SSF.format(cf_copy.numFmtId, datenum(p.v), _ssfopts);
		else p.w = SSF.format(cf_copy.numFmtId, p.v, _ssfopts);
	} catch (e) { if (opts.WTF) throw e; }
	if (!opts.cellStyles) return;
	if (cf_copy.fillId !== null) try {
		p.s = {
			fill: styles.Fills[cf_copy.fillId],
			font: styles.Fonts[cf_copy.fontId],
			border: styles.Borders[cf_copy.borderId]
		};
		if (p.s.fill.fgColor && p.s.fill.fgColor.theme && !p.s.fill.fgColor.rgb) {
			p.s.fill.fgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.fill.fgColor.theme].rgb, p.s.fill.fgColor.tint || 0);
			if (opts.WTF) p.s.fill.fgColor.raw_rgb = themes.themeElements.clrScheme[p.s.fill.fgColor.theme].rgb;
		}
		if (p.s.fill.bgColor && p.s.fill.bgColor.theme) {
			p.s.fill.bgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.fill.bgColor.theme].rgb, p.s.fill.bgColor.tint || 0);
			if (opts.WTF) p.s.fill.bgColor.raw_rgb = themes.themeElements.clrScheme[p.s.fill.bgColor.theme].rgb;
		}
		if (p.s.font.color && p.s.font.color.theme) {
			p.s.font.color.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.font.color.theme].rgb, p.s.font.color.tint || 0);
			if (opts.WTF) p.s.font.color.raw_rgb = themes.themeElements.clrScheme[p.s.font.color.theme].rgb;
		}
		if (p.s.font.color && p.s.font.color.theme) {
			p.s.font.color.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.font.color.theme].rgb, p.s.font.color.tint || 0);
			if (opts.WTF) p.s.font.color.raw_rgb = themes.themeElements.clrScheme[p.s.font.color.theme].rgb;
		}
		if (cf_copy.alignment) {
			p.s.alignment = cf_copy.alignment;
		}
	} catch (e) { if (opts.WTF && styles.Fills) throw e; }
}

function check_ws(ws/*:Worksheet*/, sname/*:string*/, i/*:number*/) {
	if(ws && ws['!ref']) {
		var range = safe_decode_range(ws['!ref']);
		if(range.e.c < range.s.c || range.e.r < range.s.r) throw new Error("Bad range (" + i + "): " + ws['!ref']);
	}
}
