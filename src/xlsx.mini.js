/*! xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
var XLSX = {};
function make_xlsx_lib(e) {
	e.version = "0.20.1";
	var r = 1200,
	t = 1252;
	var a;
	var n = [874, 932, 936, 949, 950, 1250, 1251, 1252, 1253, 1254, 1255, 1256, 1257, 1258, 1e4, ];
	var i = {
		0 : 1252,
		1 : 65001,
		2 : 65001,
		77 : 1e4,
		128 : 932,
		129 : 949,
		130 : 1361,
		134 : 936,
		136 : 950,
		161 : 1253,
		162 : 1254,
		163 : 1258,
		177 : 1255,
		178 : 1256,
		186 : 1257,
		204 : 1251,
		222 : 874,
		238 : 1250,
		255 : 1252,
		69 : 6969,
	};
	var s = function(e) {
		if (n.indexOf(e) == -1) return;
		t = i[0] = e
	};
	function l() {
		s(1252)
	}
	var o = function(e) {
		r = e;
		s(e)
	};
	function c() {
		o(1200);
		l()
	}
	function f(e) {
		var r = [];
		for (var t = 0,
		a = e.length; t < a; ++t) r[t] = e.charCodeAt(t);
		return r
	}
	function u(e) {
		var r = [];
		for (var t = 0; t < e.length >> 1; ++t) r[t] = String.fromCharCode(e.charCodeAt(2 * t) + (e.charCodeAt(2 * t + 1) << 8));
		return r.join("")
	}
	function h(e) {
		var r = [];
		for (var t = 0; t < e.length >> 1; ++t) r[t] = String.fromCharCode(e[2 * t] + (e[2 * t + 1] << 8));
		return r.join("")
	}
	function d(e) {
		var r = [];
		for (var t = 0; t < e.length >> 1; ++t) r[t] = String.fromCharCode(e.charCodeAt(2 * t + 1) + (e.charCodeAt(2 * t) << 8));
		return r.join("")
	}
	var p = function(e) {
		var r = e.charCodeAt(0),
		t = e.charCodeAt(1);
		if (r == 255 && t == 254) return u(e.slice(2));
		if (r == 254 && t == 255) return d(e.slice(2));
		if (r == 65279) return e.slice(1);
		return e
	};
	var m = function Nc(e) {
		return String.fromCharCode(e)
	};
	var v = function Ic(e) {
		return String.fromCharCode(e)
	};
	var b = null;
	var w = true;
	var k = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
	function y(e) {
		var r = "";
		var t = 0,
		a = 0,
		n = 0,
		i = 0,
		s = 0,
		l = 0,
		o = 0;
		for (var c = 0; c < e.length;) {
			t = e.charCodeAt(c++);
			i = t >> 2;
			a = e.charCodeAt(c++);
			s = ((t & 3) << 4) | (a >> 4);
			n = e.charCodeAt(c++);
			l = ((a & 15) << 2) | (n >> 6);
			o = n & 63;
			if (isNaN(a)) {
				l = o = 64
			} else if (isNaN(n)) {
				o = 64
			}
			r += k.charAt(i) + k.charAt(s) + k.charAt(l) + k.charAt(o)
		}
		return r
	}
	function C(e) {
		var r = "";
		var t = 0,
		a = 0,
		n = 0,
		i = 0,
		s = 0,
		l = 0,
		o = 0;
		e = e.replace(/^data:([^\/]+\/[^\/]+)?;base64\,/, "").replace(/[^\w\+\/\=]/g, "");
		for (var c = 0; c < e.length;) {
			i = k.indexOf(e.charAt(c++));
			s = k.indexOf(e.charAt(c++));
			t = (i << 2) | (s >> 4);
			r += String.fromCharCode(t);
			l = k.indexOf(e.charAt(c++));
			a = ((s & 15) << 4) | (l >> 2);
			if (l !== 64) {
				r += String.fromCharCode(a)
			}
			o = k.indexOf(e.charAt(c++));
			n = ((l & 3) << 6) | o;
			if (o !== 64) {
				r += String.fromCharCode(n)
			}
		}
		return r
	}
	var _ = (function() {
		return (typeof Buffer !== "undefined" && typeof undefined !== "undefined" && typeof {} !== "undefined" && !!{}.node)
	})();
	var A = (function() {
		if (typeof Buffer !== "undefined") {
			var e = !Buffer.from;
			if (!e) try {
				Buffer.from("foo", "utf8")
			} catch(r) {
				e = true
			}
			return e ?
			function(e, r) {
				return r ? new Buffer(e, r) : new Buffer(e)
			}: Buffer.from.bind(Buffer)
		}
		return function() {}
	})();
	var T = (function() {
		if (typeof Buffer === "undefined") return false;
		var e = A([65, 0]);
		if (!e) return false;
		var r = e.toString("utf16le");
		return r.length == 1
	})();
	function E(e) {
		if (_) return Buffer.alloc ? Buffer.alloc(e) : new Buffer(e);
		return typeof Uint8Array != "undefined" ? new Uint8Array(e) : new Array(e)
	}
	function F(e) {
		if (_) return Buffer.allocUnsafe ? Buffer.allocUnsafe(e) : new Buffer(e);
		return typeof Uint8Array != "undefined" ? new Uint8Array(e) : new Array(e)
	}
	var D = function Pc(e) {
		if (_) return A(e, "binary");
		return e.split("").map(function(e) {
			return e.charCodeAt(0) & 255
		})
	};
	function M(e) {
		if (Array.isArray(e)) return e.map(function(e) {
			return String.fromCharCode(e)
		}).join("");
		var r = [];
		for (var t = 0; t < e.length; ++t) r[t] = String.fromCharCode(e[t]);
		return r.join("")
	}
	function I(e) {
		if (typeof ArrayBuffer == "undefined") throw new Error("Unsupported");
		if (e instanceof ArrayBuffer) return I(new Uint8Array(e));
		var r = new Array(e.length);
		for (var t = 0; t < e.length; ++t) r[t] = e[t];
		return r
	}
	var P = _ ?
	function(e) {
		return Buffer.concat(e.map(function(e) {
			return Buffer.isBuffer(e) ? e: A(e)
		}))
	}: function(e) {
		if (typeof Uint8Array !== "undefined") {
			var r = 0,
			t = 0;
			for (r = 0; r < e.length; ++r) t += e[r].length;
			var a = new Uint8Array(t);
			var n = 0;
			for (r = 0, t = 0; r < e.length; t += n, ++r) {
				n = e[r].length;
				if (e[r] instanceof Uint8Array) a.set(e[r], t);
				else if (typeof e[r] == "string") a.set(new Uint8Array(D(e[r])), t);
				else a.set(new Uint8Array(e[r]), t)
			}
			return a
		}
		return [].concat.apply([], e.map(function(e) {
			return Array.isArray(e) ? e: [].slice.call(e)
		}))
	};
	function R(e) {
		var r = [],
		t = 0,
		a = e.length + 250;
		var n = E(e.length + 255);
		for (var i = 0; i < e.length; ++i) {
			var s = e.charCodeAt(i);
			if (s < 128) n[t++] = s;
			else if (s < 2048) {
				n[t++] = 192 | ((s >> 6) & 31);
				n[t++] = 128 | (s & 63)
			} else if (s >= 55296 && s < 57344) {
				s = (s & 1023) + 64;
				var l = e.charCodeAt(++i) & 1023;
				n[t++] = 240 | ((s >> 8) & 7);
				n[t++] = 128 | ((s >> 2) & 63);
				n[t++] = 128 | ((l >> 6) & 15) | ((s & 3) << 4);
				n[t++] = 128 | (l & 63)
			} else {
				n[t++] = 224 | ((s >> 12) & 15);
				n[t++] = 128 | ((s >> 6) & 63);
				n[t++] = 128 | (s & 63)
			}
			if (t > a) {
				r.push(n.slice(0, t));
				t = 0;
				n = E(65535);
				a = 65530
			}
		}
		r.push(n.slice(0, t));
		return P(r)
	}
	var L = /\u0000/g,
	B = /[\u0001-\u0006]/g;
	function U(e) {
		var r = "",
		t = e.length - 1;
		while (t >= 0) r += e.charAt(t--);
		return r
	}
	function z(e, r) {
		var t = "" + e;
		return t.length >= r ? t: xr("0", r - t.length) + t
	}
	function W(e, r) {
		var t = "" + e;
		return t.length >= r ? t: xr(" ", r - t.length) + t
	}
	function j(e, r) {
		var t = "" + e;
		return t.length >= r ? t: t + xr(" ", r - t.length)
	}
	function H(e, r) {
		var t = "" + Math.round(e);
		return t.length >= r ? t: xr("0", r - t.length) + t
	}
	function V(e, r) {
		var t = "" + e;
		return t.length >= r ? t: xr("0", r - t.length) + t
	}
	var X = Math.pow(2, 32);
	function G(e, r) {
		if (e > X || e < -X) return H(e, r);
		var t = Math.round(e);
		return V(t, r)
	}
	function Y(e, r) {
		r = r || 0;
		return (e.length >= 7 + r && (e.charCodeAt(r) | 32) === 103 && (e.charCodeAt(r + 1) | 32) === 101 && (e.charCodeAt(r + 2) | 32) === 110 && (e.charCodeAt(r + 3) | 32) === 101 && (e.charCodeAt(r + 4) | 32) === 114 && (e.charCodeAt(r + 5) | 32) === 97 && (e.charCodeAt(r + 6) | 32) === 108)
	}
	var J = [["Sun", "Sunday"], ["Mon", "Monday"], ["Tue", "Tuesday"], ["Wed", "Wednesday"], ["Thu", "Thursday"], ["Fri", "Friday"], ["Sat", "Saturday"], ];
	var Z = [["J", "Jan", "January"], ["F", "Feb", "February"], ["M", "Mar", "March"], ["A", "Apr", "April"], ["M", "May", "May"], ["J", "Jun", "June"], ["J", "Jul", "July"], ["A", "Aug", "August"], ["S", "Sep", "September"], ["O", "Oct", "October"], ["N", "Nov", "November"], ["D", "Dec", "December"], ];
	function K(e) {
		if (!e) e = {};
		e[0] = "General";
		e[1] = "0";
		e[2] = "0.00";
		e[3] = "#,##0";
		e[4] = "#,##0.00";
		e[9] = "0%";
		e[10] = "0.00%";
		e[11] = "0.00E+00";
		e[12] = "# ?/?";
		e[13] = "# ??/??";
		e[14] = "m/d/yy";
		e[15] = "d-mmm-yy";
		e[16] = "d-mmm";
		e[17] = "mmm-yy";
		e[18] = "h:mm AM/PM";
		e[19] = "h:mm:ss AM/PM";
		e[20] = "h:mm";
		e[21] = "h:mm:ss";
		e[22] = "m/d/yy h:mm";
		e[37] = "#,##0 ;(#,##0)";
		e[38] = "#,##0 ;[Red](#,##0)";
		e[39] = "#,##0.00;(#,##0.00)";
		e[40] = "#,##0.00;[Red](#,##0.00)";
		e[45] = "mm:ss";
		e[46] = "[h]:mm:ss";
		e[47] = "mmss.0";
		e[48] = "##0.0E+0";
		e[49] = "@";
		e[56] = '"上午/下午 "hh"時"mm"分"ss"秒 "';
		return e
	}
	var q = {
		0 : "General",
		1 : "0",
		2 : "0.00",
		3 : "#,##0",
		4 : "#,##0.00",
		9 : "0%",
		10 : "0.00%",
		11 : "0.00E+00",
		12 : "# ?/?",
		13 : "# ??/??",
		14 : "m/d/yy",
		15 : "d-mmm-yy",
		16 : "d-mmm",
		17 : "mmm-yy",
		18 : "h:mm AM/PM",
		19 : "h:mm:ss AM/PM",
		20 : "h:mm",
		21 : "h:mm:ss",
		22 : "m/d/yy h:mm",
		37 : "#,##0 ;(#,##0)",
		38 : "#,##0 ;[Red](#,##0)",
		39 : "#,##0.00;(#,##0.00)",
		40 : "#,##0.00;[Red](#,##0.00)",
		45 : "mm:ss",
		46 : "[h]:mm:ss",
		47 : "mmss.0",
		48 : "##0.0E+0",
		49 : "@",
		56 : '"上午/下午 "hh"時"mm"分"ss"秒 "',
	};
	var Q = {
		5 : 37,
		6 : 38,
		7 : 39,
		8 : 40,
		23 : 0,
		24 : 0,
		25 : 0,
		26 : 0,
		27 : 14,
		28 : 14,
		29 : 14,
		30 : 14,
		31 : 14,
		50 : 14,
		51 : 14,
		52 : 14,
		53 : 14,
		54 : 14,
		55 : 14,
		56 : 14,
		57 : 14,
		58 : 14,
		59 : 1,
		60 : 2,
		61 : 3,
		62 : 4,
		67 : 9,
		68 : 10,
		69 : 12,
		70 : 13,
		71 : 14,
		72 : 14,
		73 : 15,
		74 : 16,
		75 : 17,
		76 : 20,
		77 : 21,
		78 : 22,
		79 : 45,
		80 : 46,
		81 : 47,
		82 : 0,
	};
	var ee = {
		5 : '"$"#,##0_);\\("$"#,##0\\)',
		63 : '"$"#,##0_);\\("$"#,##0\\)',
		6 : '"$"#,##0_);[Red]\\("$"#,##0\\)',
		64 : '"$"#,##0_);[Red]\\("$"#,##0\\)',
		7 : '"$"#,##0.00_);\\("$"#,##0.00\\)',
		65 : '"$"#,##0.00_);\\("$"#,##0.00\\)',
		8 : '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
		66 : '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
		41 : '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
		42 : '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
		43 : '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
		44 : '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
	};
	function re(e, r, t) {
		var a = e < 0 ? -1 : 1;
		var n = e * a;
		var i = 0,
		s = 1,
		l = 0;
		var o = 1,
		c = 0,
		f = 0;
		var u = Math.floor(n);
		while (c < r) {
			u = Math.floor(n);
			l = u * s + i;
			f = u * c + o;
			if (n - u < 5e-8) break;
			n = 1 / (n - u);
			i = s;
			s = l;
			o = c;
			c = f
		}
		if (f > r) {
			if (c > r) {
				f = o;
				l = i
			} else {
				f = c;
				l = s
			}
		}
		if (!t) return [0, a * l, f];
		var h = Math.floor((a * l) / f);
		return [h, a * l - h * f, f]
	}
	function te(e) {
		var r = e.toPrecision(16);
		if (r.indexOf("e") > -1) {
			var t = r.slice(0, r.indexOf("e"));
			t = t.indexOf(".") > -1 ? t.slice(0, t.slice(0, 2) == "0." ? 17 : 16) : t.slice(0, 15) + xr("0", t.length - 15);
			return t + r.slice(r.indexOf("e"))
		}
		var a = r.indexOf(".") > -1 ? r.slice(0, r.slice(0, 2) == "0." ? 17 : 16) : r.slice(0, 15) + xr("0", r.length - 15);
		return Number(a)
	}
	function ae(e, r, t) {
		if (e > 2958465 || e < 0) return null;
		e = te(e);
		var a = e | 0,
		n = Math.floor(86400 * (e - a)),
		i = 0;
		var s = [];
		var l = {
			D: a,
			T: n,
			u: 86400 * (e - a) - n,
			y: 0,
			m: 0,
			d: 0,
			H: 0,
			M: 0,
			S: 0,
			q: 0,
		};
		if (Math.abs(l.u) < 1e-6) l.u = 0;
		if (r && r.date1904) a += 1462;
		if (l.u > 0.9999) {
			l.u = 0;
			if (++n == 86400) {
				l.T = n = 0; ++a; ++l.D
			}
		}
		if (a === 60) {
			s = t ? [1317, 10, 29] : [1900, 2, 29];
			i = 3
		} else if (a === 0) {
			s = t ? [1317, 8, 29] : [1900, 1, 0];
			i = 6
		} else {
			if (a > 60)--a;
			var o = new Date(1900, 0, 1);
			o.setDate(o.getDate() + a - 1);
			s = [o.getFullYear(), o.getMonth() + 1, o.getDate()];
			i = o.getDay();
			if (a < 60) i = (i + 6) % 7;
			if (t) i = fe(o, s)
		}
		l.y = s[0];
		l.m = s[1];
		l.d = s[2];
		l.S = n % 60;
		n = Math.floor(n / 60);
		l.M = n % 60;
		n = Math.floor(n / 60);
		l.H = n;
		l.q = i;
		return l
	}
	function ne(e) {
		return e.indexOf(".") == -1 ? e: e.replace(/(?:\.0*|(\.\d*[1-9])0+)$/, "$1")
	}
	function ie(e) {
		if (e.indexOf("E") == -1) return e;
		return e.replace(/(?:\.0*|(\.\d*[1-9])0+)[Ee]/, "$1E").replace(/(E[+-])(\d)$/, "$10$2")
	}
	function se(e) {
		var r = e < 0 ? 12 : 11;
		var t = ne(e.toFixed(12));
		if (t.length <= r) return t;
		t = e.toPrecision(10);
		if (t.length <= r) return t;
		return e.toExponential(5)
	}
	function le(e) {
		var r = ne(e.toFixed(11));
		return r.length > (e < 0 ? 12 : 11) || r === "0" || r === "-0" ? e.toPrecision(6) : r
	}
	function oe(e) {
		var r = Math.floor(Math.log(Math.abs(e)) * Math.LOG10E),
		t;
		if (r >= -4 && r <= -1) t = e.toPrecision(10 + r);
		else if (Math.abs(r) <= 9) t = se(e);
		else if (r === 10) t = e.toFixed(10).substr(0, 12);
		else t = le(e);
		return ne(ie(t.toUpperCase()))
	}
	function ce(e, r) {
		switch (typeof e) {
		case "string":
			return e;
		case "boolean":
			return e ? "TRUE": "FALSE";
		case "number":
			return (e | 0) === e ? e.toString(10) : oe(e);
		case "undefined":
			return "";
		case "object":
			if (e == null) return "";
			if (e instanceof Date) return $e(14, dr(e, r && r.date1904), r)
		}
		throw new Error("unsupported value in General format: " + e);
	}
	function fe(e, r) {
		r[0] -= 581;
		var t = e.getDay();
		if (e < 60) t = (t + 6) % 7;
		return t
	}
	function ue(e, r, t, a) {
		var n = "",
		i = 0,
		s = 0,
		l = t.y,
		o, c = 0;
		switch (e) {
		case 98:
			l = t.y + 543;
		case 121:
			switch (r.length) {
			case 1:
			case 2:
				o = l % 100;
				c = 2;
				break;
			default:
				o = l % 1e4;
				c = 4;
				break
			}
			break;
		case 109:
			switch (r.length) {
			case 1:
			case 2:
				o = t.m;
				c = r.length;
				break;
			case 3:
				return Z[t.m - 1][1];
			case 5:
				return Z[t.m - 1][0];
			default:
				return Z[t.m - 1][2]
			}
			break;
		case 100:
			switch (r.length) {
			case 1:
			case 2:
				o = t.d;
				c = r.length;
				break;
			case 3:
				return J[t.q][0];
			default:
				return J[t.q][1]
			}
			break;
		case 104:
			switch (r.length) {
			case 1:
			case 2:
				o = 1 + ((t.H + 11) % 12);
				c = r.length;
				break;
			default:
				throw "bad hour format: " + r;
			}
			break;
		case 72:
			switch (r.length) {
			case 1:
			case 2:
				o = t.H;
				c = r.length;
				break;
			default:
				throw "bad hour format: " + r;
			}
			break;
		case 77:
			switch (r.length) {
			case 1:
			case 2:
				o = t.M;
				c = r.length;
				break;
			default:
				throw "bad minute format: " + r;
			}
			break;
		case 115:
			if (r != "s" && r != "ss" && r != ".0" && r != ".00" && r != ".000") throw "bad second format: " + r;
			if (t.u === 0 && (r == "s" || r == "ss")) return z(t.S, r.length);
			if (a >= 2) s = a === 3 ? 1e3: 100;
			else s = a === 1 ? 10 : 1;
			i = Math.round(s * (t.S + t.u));
			if (i >= 60 * s) i = 0;
			if (r === "s") return i === 0 ? "0": "" + i / s;
			n = z(i, 2 + a);
			if (r === "ss") return n.substr(0, 2);
			return "." + n.substr(2, r.length - 1);
		case 90:
			switch (r) {
			case "[h]":
			case "[hh]":
				o = t.D * 24 + t.H;
				break;
			case "[m]":
			case "[mm]":
				o = (t.D * 24 + t.H) * 60 + t.M;
				break;
			case "[s]":
			case "[ss]":
				o = ((t.D * 24 + t.H) * 60 + t.M) * 60 + (a == 0 ? Math.round(t.S + t.u) : t.S);
				break;
			default:
				throw "bad abstime format: " + r;
			}
			c = r.length === 3 ? 1 : 2;
			break;
		case 101:
			o = l;
			c = 1;
			break
		}
		var f = c > 0 ? z(o, c) : "";
		return f
	}
	function he(e) {
		var r = 3;
		if (e.length <= r) return e;
		var t = e.length % r,
		a = e.substr(0, t);
		for (; t != e.length; t += r) a += (a.length > 0 ? ",": "") + e.substr(t, r);
		return a
	}
	var de = /%/g;
	function pe(e, r, t) {
		var a = r.replace(de, ""),
		n = r.length - a.length;
		return Ne(e, a, t * Math.pow(10, 2 * n)) + xr("%", n)
	}
	function me(e, r, t) {
		var a = r.length - 1;
		while (r.charCodeAt(a - 1) === 44)--a;
		return Ne(e, r.substr(0, a), t / Math.pow(10, 3 * (r.length - a)))
	}
	function ve(e, r) {
		var t;
		var a = e.indexOf("E") - e.indexOf(".") - 1;
		if (e.match(/^#+0.0E\+0$/)) {
			if (r == 0) return "0.0E+0";
			else if (r < 0) return "-" + ve(e, -r);
			var n = e.indexOf(".");
			if (n === -1) n = e.indexOf("E");
			var i = Math.floor(Math.log(r) * Math.LOG10E) % n;
			if (i < 0) i += n;
			t = (r / Math.pow(10, i)).toPrecision(a + 1 + ((n + i) % n));
			if (t.indexOf("e") === -1) {
				var s = Math.floor(Math.log(r) * Math.LOG10E);
				if (t.indexOf(".") === -1) t = t.charAt(0) + "." + t.substr(1) + "E+" + (s - t.length + i);
				else t += "E+" + (s - i);
				while (t.substr(0, 2) === "0.") {
					t = t.charAt(0) + t.substr(2, n) + "." + t.substr(2 + n);
					t = t.replace(/^0+([1-9])/, "$1").replace(/^0+\./, "0.")
				}
				t = t.replace(/\+-/, "-")
			}
			t = t.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/,
			function(e, r, t, a) {
				return r + t + a.substr(0, (n + i) % n) + "." + a.substr(i) + "E"
			})
		} else t = r.toExponential(a);
		if (e.match(/E\+00$/) && t.match(/e[+-]\d$/)) t = t.substr(0, t.length - 1) + "0" + t.charAt(t.length - 1);
		if (e.match(/E\-/) && t.match(/e\+/)) t = t.replace(/e\+/, "e");
		return t.replace("e", "E")
	}
	var ge = /# (\?+)( ?)\/( ?)(\d+)/;
	function be(e, r, t) {
		var a = parseInt(e[4], 10),
		n = Math.round(r * a),
		i = Math.floor(n / a);
		var s = n - i * a,
		l = a;
		return (t + (i === 0 ? "": "" + i) + " " + (s === 0 ? xr(" ", e[1].length + 1 + e[4].length) : W(s, e[1].length) + e[2] + "/" + e[3] + z(l, e[4].length)))
	}
	function we(e, r, t) {
		return t + (r === 0 ? "": "" + r) + xr(" ", e[1].length + 2 + e[4].length)
	}
	var ke = /^#*0*\.([0#]+)/;
	var ye = /\).*[0#]/;
	var xe = /\(###\) ###\\?-####/;
	function Se(e) {
		var r = "",
		t;
		for (var a = 0; a != e.length; ++a) switch ((t = e.charCodeAt(a))) {
		case 35:
			break;
		case 63:
			r += " ";
			break;
		case 48:
			r += "0";
			break;
		default:
			r += String.fromCharCode(t)
		}
		return r
	}
	function Ce(e, r) {
		var t = Math.pow(10, r);
		return "" + Math.round(e * t) / t
	}
	function _e(e, r) {
		var t = e - Math.floor(e),
		a = Math.pow(10, r);
		if (r < ("" + Math.round(t * a)).length) return 0;
		return Math.round(t * a)
	}
	function Ae(e, r) {
		if (r < ("" + Math.round((e - Math.floor(e)) * Math.pow(10, r))).length) {
			return 1
		}
		return 0
	}
	function Te(e) {
		if (e < 2147483647 && e > -2147483648) return "" + (e >= 0 ? e | 0 : (e - 1) | 0);
		return "" + Math.floor(e)
	}
	function Ee(e, r, t) {
		if (e.charCodeAt(0) === 40 && !r.match(ye)) {
			var a = r.replace(/\( */, "").replace(/ \)/, "").replace(/\)/, "");
			if (t >= 0) return Ee("n", a, t);
			return "(" + Ee("n", a, -t) + ")"
		}
		if (r.charCodeAt(r.length - 1) === 44) return me(e, r, t);
		if (r.indexOf("%") !== -1) return pe(e, r, t);
		if (r.indexOf("E") !== -1) return ve(r, t);
		if (r.charCodeAt(0) === 36) return "$" + Ee(e, r.substr(r.charAt(1) == " " ? 2 : 1), t);
		var n;
		var i, s, l, o = Math.abs(t),
		c = t < 0 ? "-": "";
		if (r.match(/^00+$/)) return c + G(o, r.length);
		if (r.match(/^[#?]+$/)) {
			n = G(t, 0);
			if (n === "0") n = "";
			return n.length > r.length ? n: Se(r.substr(0, r.length - n.length)) + n
		}
		if ((i = r.match(ge))) return be(i, o, c);
		if (r.match(/^#+0+$/)) return c + G(o, r.length - r.indexOf("0"));
		if ((i = r.match(ke))) {
			n = Ce(t, i[1].length).replace(/^([^\.]+)$/, "$1." + Se(i[1])).replace(/\.$/, "." + Se(i[1])).replace(/\.(\d*)$/,
			function(e, r) {
				return "." + r + xr("0", Se(i[1]).length - r.length)
			});
			return r.indexOf("0.") !== -1 ? n: n.replace(/^0\./, ".")
		}
		r = r.replace(/^#+([0.])/, "$1");
		if ((i = r.match(/^(0*)\.(#*)$/))) {
			return (c + Ce(o, i[2].length).replace(/\.(\d*[1-9])0*$/, ".$1").replace(/^(-?\d*)$/, "$1.").replace(/^0\./, i[1].length ? "0.": "."))
		}
		if ((i = r.match(/^#{1,3},##0(\.?)$/))) return c + he(G(o, 0));
		if ((i = r.match(/^#,##0\.([#0]*0)$/))) {
			return t < 0 ? "-" + Ee(e, r, -t) : he("" + (Math.floor(t) + Ae(t, i[1].length))) + "." + z(_e(t, i[1].length), i[1].length)
		}
		if ((i = r.match(/^#,#*,#0/))) return Ee(e, r.replace(/^#,#*,/, ""), t);
		if ((i = r.match(/^([0#]+)(\\?-([0#]+))+$/))) {
			n = U(Ee(e, r.replace(/[\\-]/g, ""), t));
			s = 0;
			return U(U(r.replace(/\\/g, "")).replace(/[0#]/g,
			function(e) {
				return s < n.length ? n.charAt(s++) : e === "0" ? "0": ""
			}))
		}
		if (r.match(xe)) {
			n = Ee(e, "##########", t);
			return "(" + n.substr(0, 3) + ") " + n.substr(3, 3) + "-" + n.substr(6)
		}
		var f = "";
		if ((i = r.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
			s = Math.min(i[4].length, 7);
			l = re(o, Math.pow(10, s) - 1, false);
			n = "" + c;
			f = Ne("n", i[1], l[1]);
			if (f.charAt(f.length - 1) == " ") f = f.substr(0, f.length - 1) + "0";
			n += f + i[2] + "/" + i[3];
			f = j(l[2], s);
			if (f.length < i[4].length) f = Se(i[4].substr(i[4].length - f.length)) + f;
			n += f;
			return n
		}
		if ((i = r.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
			s = Math.min(Math.max(i[1].length, i[4].length), 7);
			l = re(o, Math.pow(10, s) - 1, true);
			return (c + (l[0] || (l[1] ? "": "0")) + " " + (l[1] ? W(l[1], s) + i[2] + "/" + i[3] + j(l[2], s) : xr(" ", 2 * s + 1 + i[2].length + i[3].length)))
		}
		if ((i = r.match(/^[#0?]+$/))) {
			n = G(t, 0);
			if (r.length <= n.length) return n;
			return Se(r.substr(0, r.length - n.length)) + n
		}
		if ((i = r.match(/^([#0?]+)\.([#0]+)$/))) {
			n = "" + t.toFixed(Math.min(i[2].length, 10)).replace(/([^0])0+$/, "$1");
			s = n.indexOf(".");
			var u = r.indexOf(".") - s,
			h = r.length - n.length - u;
			return Se(r.substr(0, u) + n + r.substr(r.length - h))
		}
		if ((i = r.match(/^00,000\.([#0]*0)$/))) {
			s = _e(t, i[1].length);
			return t < 0 ? "-" + Ee(e, r, -t) : he(Te(t)).replace(/^\d,\d{3}$/, "0$&").replace(/^\d*$/,
			function(e) {
				return "00," + (e.length < 3 ? z(0, 3 - e.length) : "") + e
			}) + "." + z(s, i[1].length)
		}
		switch (r) {
		case "###,##0.00":
			return Ee(e, "#,##0.00", t);
		case "###,###":
		case "##,###":
		case "#,###":
			var d = he(G(o, 0));
			return d !== "0" ? c + d: "";
		case "###,###.00":
			return Ee(e, "###,##0.00", t).replace(/^0\./, ".");
		case "#,###.00":
			return Ee(e, "#,##0.00", t).replace(/^0\./, ".");
		default:
		}
		throw new Error("unsupported format |" + r + "|");
	}
	function Fe(e, r, t) {
		var a = r.length - 1;
		while (r.charCodeAt(a - 1) === 44)--a;
		return Ne(e, r.substr(0, a), t / Math.pow(10, 3 * (r.length - a)))
	}
	function De(e, r, t) {
		var a = r.replace(de, ""),
		n = r.length - a.length;
		return Ne(e, a, t * Math.pow(10, 2 * n)) + xr("%", n)
	}
	function Oe(e, r) {
		var t;
		var a = e.indexOf("E") - e.indexOf(".") - 1;
		if (e.match(/^#+0.0E\+0$/)) {
			if (r == 0) return "0.0E+0";
			else if (r < 0) return "-" + Oe(e, -r);
			var n = e.indexOf(".");
			if (n === -1) n = e.indexOf("E");
			var i = Math.floor(Math.log(r) * Math.LOG10E) % n;
			if (i < 0) i += n;
			t = (r / Math.pow(10, i)).toPrecision(a + 1 + ((n + i) % n));
			if (!t.match(/[Ee]/)) {
				var s = Math.floor(Math.log(r) * Math.LOG10E);
				if (t.indexOf(".") === -1) t = t.charAt(0) + "." + t.substr(1) + "E+" + (s - t.length + i);
				else t += "E+" + (s - i);
				t = t.replace(/\+-/, "-")
			}
			t = t.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/,
			function(e, r, t, a) {
				return r + t + a.substr(0, (n + i) % n) + "." + a.substr(i) + "E"
			})
		} else t = r.toExponential(a);
		if (e.match(/E\+00$/) && t.match(/e[+-]\d$/)) t = t.substr(0, t.length - 1) + "0" + t.charAt(t.length - 1);
		if (e.match(/E\-/) && t.match(/e\+/)) t = t.replace(/e\+/, "e");
		return t.replace("e", "E")
	}
	function Me(e, r, t) {
		if (e.charCodeAt(0) === 40 && !r.match(ye)) {
			var a = r.replace(/\( */, "").replace(/ \)/, "").replace(/\)/, "");
			if (t >= 0) return Me("n", a, t);
			return "(" + Me("n", a, -t) + ")"
		}
		if (r.charCodeAt(r.length - 1) === 44) return Fe(e, r, t);
		if (r.indexOf("%") !== -1) return De(e, r, t);
		if (r.indexOf("E") !== -1) return Oe(r, t);
		if (r.charCodeAt(0) === 36) return "$" + Me(e, r.substr(r.charAt(1) == " " ? 2 : 1), t);
		var n;
		var i, s, l, o = Math.abs(t),
		c = t < 0 ? "-": "";
		if (r.match(/^00+$/)) return c + z(o, r.length);
		if (r.match(/^[#?]+$/)) {
			n = "" + t;
			if (t === 0) n = "";
			return n.length > r.length ? n: Se(r.substr(0, r.length - n.length)) + n
		}
		if ((i = r.match(ge))) return we(i, o, c);
		if (r.match(/^#+0+$/)) return c + z(o, r.length - r.indexOf("0"));
		if ((i = r.match(ke))) {
			n = ("" + t).replace(/^([^\.]+)$/, "$1." + Se(i[1])).replace(/\.$/, "." + Se(i[1]));
			n = n.replace(/\.(\d*)$/,
			function(e, r) {
				return "." + r + xr("0", Se(i[1]).length - r.length)
			});
			return r.indexOf("0.") !== -1 ? n: n.replace(/^0\./, ".")
		}
		r = r.replace(/^#+([0.])/, "$1");
		if ((i = r.match(/^(0*)\.(#*)$/))) {
			return (c + ("" + o).replace(/\.(\d*[1-9])0*$/, ".$1").replace(/^(-?\d*)$/, "$1.").replace(/^0\./, i[1].length ? "0.": "."))
		}
		if ((i = r.match(/^#{1,3},##0(\.?)$/))) return c + he("" + o);
		if ((i = r.match(/^#,##0\.([#0]*0)$/))) {
			return t < 0 ? "-" + Me(e, r, -t) : he("" + t) + "." + xr("0", i[1].length)
		}
		if ((i = r.match(/^#,#*,#0/))) return Me(e, r.replace(/^#,#*,/, ""), t);
		if ((i = r.match(/^([0#]+)(\\?-([0#]+))+$/))) {
			n = U(Me(e, r.replace(/[\\-]/g, ""), t));
			s = 0;
			return U(U(r.replace(/\\/g, "")).replace(/[0#]/g,
			function(e) {
				return s < n.length ? n.charAt(s++) : e === "0" ? "0": ""
			}))
		}
		if (r.match(xe)) {
			n = Me(e, "##########", t);
			return "(" + n.substr(0, 3) + ") " + n.substr(3, 3) + "-" + n.substr(6)
		}
		var f = "";
		if ((i = r.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
			s = Math.min(i[4].length, 7);
			l = re(o, Math.pow(10, s) - 1, false);
			n = "" + c;
			f = Ne("n", i[1], l[1]);
			if (f.charAt(f.length - 1) == " ") f = f.substr(0, f.length - 1) + "0";
			n += f + i[2] + "/" + i[3];
			f = j(l[2], s);
			if (f.length < i[4].length) f = Se(i[4].substr(i[4].length - f.length)) + f;
			n += f;
			return n
		}
		if ((i = r.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
			s = Math.min(Math.max(i[1].length, i[4].length), 7);
			l = re(o, Math.pow(10, s) - 1, true);
			return (c + (l[0] || (l[1] ? "": "0")) + " " + (l[1] ? W(l[1], s) + i[2] + "/" + i[3] + j(l[2], s) : xr(" ", 2 * s + 1 + i[2].length + i[3].length)))
		}
		if ((i = r.match(/^[#0?]+$/))) {
			n = "" + t;
			if (r.length <= n.length) return n;
			return Se(r.substr(0, r.length - n.length)) + n
		}
		if ((i = r.match(/^([#0]+)\.([#0]+)$/))) {
			n = "" + t.toFixed(Math.min(i[2].length, 10)).replace(/([^0])0+$/, "$1");
			s = n.indexOf(".");
			var u = r.indexOf(".") - s,
			h = r.length - n.length - u;
			return Se(r.substr(0, u) + n + r.substr(r.length - h))
		}
		if ((i = r.match(/^00,000\.([#0]*0)$/))) {
			return t < 0 ? "-" + Me(e, r, -t) : he("" + t).replace(/^\d,\d{3}$/, "0$&").replace(/^\d*$/,
			function(e) {
				return "00," + (e.length < 3 ? z(0, 3 - e.length) : "") + e
			}) + "." + z(0, i[1].length)
		}
		switch (r) {
		case "###,###":
		case "##,###":
		case "#,###":
			var d = he("" + o);
			return d !== "0" ? c + d: "";
		default:
			if (r.match(/\.[0#?]*$/)) return (Me(e, r.slice(0, r.lastIndexOf(".")), t) + Se(r.slice(r.lastIndexOf("."))))
		}
		throw new Error("unsupported format |" + r + "|");
	}
	function Ne(e, r, t) {
		return (t | 0) === t ? Me(e, r, t) : Ee(e, r, t)
	}
	function Ie(e) {
		var r = [];
		var t = false;
		for (var a = 0,
		n = 0; a < e.length; ++a) switch (e.charCodeAt(a)) {
		case 34:
			t = !t;
			break;
		case 95:
		case 42:
		case 92:
			++a;
			break;
		case 59:
			r[r.length] = e.substr(n, a - n);
			n = a + 1
		}
		r[r.length] = e.substr(n);
		if (t === true) throw new Error("Format |" + e + "| unterminated string ");
		return r
	}
	var Pe = /\[[HhMmSs\u0E0A\u0E19\u0E17]*\]/;
	function Re(e) {
		var r = 0,
		t = "",
		a = "";
		while (r < e.length) {
			switch ((t = e.charAt(r))) {
			case "G":
				if (Y(e, r)) r += 6;
				r++;
				break;
			case '"':
				for (; e.charCodeAt(++r) !== 34 && r < e.length;) {}++r;
				break;
			case "\\":
				r += 2;
				break;
			case "_":
				r += 2;
				break;
			case "@":
				++r;
				break;
			case "B":
			case "b":
				if (e.charAt(r + 1) === "1" || e.charAt(r + 1) === "2") return true;
			case "M":
			case "D":
			case "Y":
			case "H":
			case "S":
			case "E":
			case "m":
			case "d":
			case "y":
			case "h":
			case "s":
			case "e":
			case "g":
				return true;
			case "A":
			case "a":
			case "上":
				if (e.substr(r, 3).toUpperCase() === "A/P") return true;
				if (e.substr(r, 5).toUpperCase() === "AM/PM") return true;
				if (e.substr(r, 5).toUpperCase() === "上午/下午") return true; ++r;
				break;
			case "[":
				a = t;
				while (e.charAt(r++) !== "]" && r < e.length) a += e.charAt(r);
				if (a.match(Pe)) return true;
				break;
			case ".":
			case "0":
			case "#":
				while (r < e.length && ("0#?.,E+-%".indexOf((t = e.charAt(++r))) > -1 || (t == "\\" && e.charAt(r + 1) == "-" && "0#".indexOf(e.charAt(r + 2)) > -1))) {}
				break;
			case "?":
				while (e.charAt(++r) === t) {}
				break;
			case "*":
				++r;
				if (e.charAt(r) == " " || e.charAt(r) == "*")++r;
				break;
			case "(":
			case ")":
				++r;
				break;
			case "1":
			case "2":
			case "3":
			case "4":
			case "5":
			case "6":
			case "7":
			case "8":
			case "9":
				while (r < e.length && "0123456789".indexOf(e.charAt(++r)) > -1) {}
				break;
			case " ":
				++r;
				break;
			default:
				++r;
				break
			}
		}
		return false
	}
	function Le(e, r, t, a) {
		var n = [],
		i = "",
		s = 0,
		l = "",
		o = "t",
		c,
		f,
		u;
		var h = "H";
		while (s < e.length) {
			switch ((l = e.charAt(s))) {
			case "G":
				if (!Y(e, s)) throw new Error("unrecognized character " + l + " in " + e);
				n[n.length] = {
					t: "G",
					v: "General",
				};
				s += 7;
				break;
			case '"':
				for (i = ""; (u = e.charCodeAt(++s)) !== 34 && s < e.length;) i += String.fromCharCode(u);
				n[n.length] = {
					t: "t",
					v: i,
				}; ++s;
				break;
			case "\\":
				var d = e.charAt(++s),
				p = d === "(" || d === ")" ? d: "t";
				n[n.length] = {
					t: p,
					v: d,
				}; ++s;
				break;
			case "_":
				n[n.length] = {
					t: "t",
					v: " ",
				};
				s += 2;
				break;
			case "@":
				n[n.length] = {
					t: "T",
					v: r,
				}; ++s;
				break;
			case "B":
			case "b":
				if (e.charAt(s + 1) === "1" || e.charAt(s + 1) === "2") {
					if (c == null) {
						c = ae(r, t, e.charAt(s + 1) === "2");
						if (c == null) return ""
					}
					n[n.length] = {
						t: "X",
						v: e.substr(s, 2),
					};
					o = l;
					s += 2;
					break
				}
			case "M":
			case "D":
			case "Y":
			case "H":
			case "S":
			case "E":
				l = l.toLowerCase();
			case "m":
			case "d":
			case "y":
			case "h":
			case "s":
			case "e":
			case "g":
				if (r < 0) return "";
				if (c == null) {
					c = ae(r, t);
					if (c == null) return ""
				}
				i = l;
				while (++s < e.length && e.charAt(s).toLowerCase() === l) i += l;
				if (l === "m" && o.toLowerCase() === "h") l = "M";
				if (l === "h") l = h;
				n[n.length] = {
					t: l,
					v: i,
				};
				o = l;
				break;
			case "A":
			case "a":
			case "上":
				var m = {
					t: l,
					v: l,
				};
				if (c == null) c = ae(r, t);
				if (e.substr(s, 3).toUpperCase() === "A/P") {
					if (c != null) m.v = c.H >= 12 ? e.charAt(s + 2) : l;
					m.t = "T";
					h = "h";
					s += 3
				} else if (e.substr(s, 5).toUpperCase() === "AM/PM") {
					if (c != null) m.v = c.H >= 12 ? "PM": "AM";
					m.t = "T";
					s += 5;
					h = "h"
				} else if (e.substr(s, 5).toUpperCase() === "上午/下午") {
					if (c != null) m.v = c.H >= 12 ? "下午": "上午";
					m.t = "T";
					s += 5;
					h = "h"
				} else {
					m.t = "t"; ++s
				}
				if (c == null && m.t === "T") return "";
				n[n.length] = m;
				o = l;
				break;
			case "[":
				i = l;
				while (e.charAt(s++) !== "]" && s < e.length) i += e.charAt(s);
				if (i.slice(-1) !== "]") throw 'unterminated "[" block: |' + i + "|";
				if (i.match(Pe)) {
					if (c == null) {
						c = ae(r, t);
						if (c == null) return ""
					}
					n[n.length] = {
						t: "Z",
						v: i.toLowerCase(),
					};
					o = i.charAt(1)
				} else if (i.indexOf("$") > -1) {
					i = (i.match(/\$([^-\[\]]*)/) || [])[1] || "$";
					if (!Re(e)) n[n.length] = {
						t: "t",
						v: i,
					}
				}
				break;
			case ".":
				if (c != null) {
					i = l;
					while (++s < e.length && (l = e.charAt(s)) === "0") i += l;
					n[n.length] = {
						t: "s",
						v: i,
					};
					break
				}
			case "0":
			case "#":
				i = l;
				while (++s < e.length && "0#?.,E+-%".indexOf((l = e.charAt(s))) > -1) i += l;
				n[n.length] = {
					t: "n",
					v: i,
				};
				break;
			case "?":
				i = l;
				while (e.charAt(++s) === l) i += l;
				n[n.length] = {
					t: l,
					v: i,
				};
				o = l;
				break;
			case "*":
				++s;
				if (e.charAt(s) == " " || e.charAt(s) == "*")++s;
				break;
			case "(":
			case ")":
				n[n.length] = {
					t: a === 1 ? "t": l,
					v: l,
				}; ++s;
				break;
			case "1":
			case "2":
			case "3":
			case "4":
			case "5":
			case "6":
			case "7":
			case "8":
			case "9":
				i = l;
				while (s < e.length && "0123456789".indexOf(e.charAt(++s)) > -1) i += e.charAt(s);
				n[n.length] = {
					t: "D",
					v: i,
				};
				break;
			case " ":
				n[n.length] = {
					t: l,
					v: l,
				}; ++s;
				break;
			case "$":
				n[n.length] = {
					t: "t",
					v: "$",
				}; ++s;
				break;
			default:
				if (",$-+/():!^&'~{}<>=€acfijklopqrtuvwxzP".indexOf(l) === -1) throw new Error("unrecognized character " + l + " in " + e);
				n[n.length] = {
					t: "t",
					v: l,
				}; ++s;
				break
			}
		}
		var v = 0,
		g = 0,
		b;
		for (s = n.length - 1, o = "t"; s >= 0; --s) {
			switch (n[s].t) {
			case "h":
			case "H":
				n[s].t = h;
				o = "h";
				if (v < 1) v = 1;
				break;
			case "s":
				if ((b = n[s].v.match(/\.0+$/))) {
					g = Math.max(g, b[0].length - 1);
					v = 4
				}
				if (v < 3) v = 3;
			case "d":
			case "y":
			case "e":
				o = n[s].t;
				break;
			case "M":
				o = n[s].t;
				if (v < 2) v = 2;
				break;
			case "m":
				if (o === "s") {
					n[s].t = "M";
					if (v < 2) v = 2
				}
				break;
			case "X":
				break;
			case "Z":
				if (v < 1 && n[s].v.match(/[Hh]/)) v = 1;
				if (v < 2 && n[s].v.match(/[Mm]/)) v = 2;
				if (v < 3 && n[s].v.match(/[Ss]/)) v = 3
			}
		}
		var w;
		switch (v) {
		case 0:
			break;
		case 1:
		case 2:
		case 3:
			if (c.u >= 0.5) {
				c.u = 0; ++c.S
			}
			if (c.S >= 60) {
				c.S = 0; ++c.M
			}
			if (c.M >= 60) {
				c.M = 0; ++c.H
			}
			if (c.H >= 24) {
				c.H = 0; ++c.D;
				w = ae(c.D);
				w.u = c.u;
				w.S = c.S;
				w.M = c.M;
				w.H = c.H;
				c = w
			}
			break;
		case 4:
			switch (g) {
			case 1:
				c.u = Math.round(c.u * 10) / 10;
				break;
			case 2:
				c.u = Math.round(c.u * 100) / 100;
				break;
			case 3:
				c.u = Math.round(c.u * 1e3) / 1e3;
				break
			}
			if (c.u >= 1) {
				c.u = 0; ++c.S
			}
			if (c.S >= 60) {
				c.S = 0; ++c.M
			}
			if (c.M >= 60) {
				c.M = 0; ++c.H
			}
			if (c.H >= 24) {
				c.H = 0; ++c.D;
				w = ae(c.D);
				w.u = c.u;
				w.S = c.S;
				w.M = c.M;
				w.H = c.H;
				c = w
			}
			break
		}
		var k = "",
		y;
		for (s = 0; s < n.length; ++s) {
			switch (n[s].t) {
			case "t":
			case "T":
			case " ":
			case "D":
				break;
			case "X":
				n[s].v = "";
				n[s].t = ";";
				break;
			case "d":
			case "m":
			case "y":
			case "h":
			case "H":
			case "M":
			case "s":
			case "e":
			case "b":
			case "Z":
				n[s].v = ue(n[s].t.charCodeAt(0), n[s].v, c, g);
				n[s].t = "t";
				break;
			case "n":
			case "?":
				y = s + 1;
				while (n[y] != null && ((l = n[y].t) === "?" || l === "D" || ((l === " " || l === "t") && n[y + 1] != null && (n[y + 1].t === "?" || (n[y + 1].t === "t" && n[y + 1].v === "/"))) || (n[s].t === "(" && (l === " " || l === "n" || l === ")")) || (l === "t" && (n[y].v === "/" || (n[y].v === " " && n[y + 1] != null && n[y + 1].t == "?"))))) {
					n[s].v += n[y].v;
					n[y] = {
						v: "",
						t: ";",
					}; ++y
				}
				k += n[s].v;
				s = y - 1;
				break;
			case "G":
				n[s].t = "t";
				n[s].v = ce(r, t);
				break
			}
		}
		var x = "",
		S, C;
		if (k.length > 0) {
			if (k.charCodeAt(0) == 40) {
				S = r < 0 && k.charCodeAt(0) === 45 ? -r: r;
				C = Ne("n", k, S)
			} else {
				S = r < 0 && a > 1 ? -r: r;
				C = Ne("n", k, S);
				if (S < 0 && n[0] && n[0].t == "t") {
					C = C.substr(1);
					n[0].v = "-" + n[0].v
				}
			}
			y = C.length - 1;
			var _ = n.length;
			for (s = 0; s < n.length; ++s) if (n[s] != null && n[s].t != "t" && n[s].v.indexOf(".") > -1) {
				_ = s;
				break
			}
			var A = n.length;
			if (_ === n.length && C.indexOf("E") === -1) {
				for (s = n.length - 1; s >= 0; --s) {
					if (n[s] == null || "n?".indexOf(n[s].t) === -1) continue;
					if (y >= n[s].v.length - 1) {
						y -= n[s].v.length;
						n[s].v = C.substr(y + 1, n[s].v.length)
					} else if (y < 0) n[s].v = "";
					else {
						n[s].v = C.substr(0, y + 1);
						y = -1
					}
					n[s].t = "t";
					A = s
				}
				if (y >= 0 && A < n.length) n[A].v = C.substr(0, y + 1) + n[A].v
			} else if (_ !== n.length && C.indexOf("E") === -1) {
				y = C.indexOf(".") - 1;
				for (s = _; s >= 0; --s) {
					if (n[s] == null || "n?".indexOf(n[s].t) === -1) continue;
					f = n[s].v.indexOf(".") > -1 && s === _ ? n[s].v.indexOf(".") - 1 : n[s].v.length - 1;
					x = n[s].v.substr(f + 1);
					for (; f >= 0; --f) {
						if (y >= 0 && (n[s].v.charAt(f) === "0" || n[s].v.charAt(f) === "#")) x = C.charAt(y--) + x
					}
					n[s].v = x;
					n[s].t = "t";
					A = s
				}
				if (y >= 0 && A < n.length) n[A].v = C.substr(0, y + 1) + n[A].v;
				y = C.indexOf(".") + 1;
				for (s = _; s < n.length; ++s) {
					if (n[s] == null || ("n?(".indexOf(n[s].t) === -1 && s !== _)) continue;
					f = n[s].v.indexOf(".") > -1 && s === _ ? n[s].v.indexOf(".") + 1 : 0;
					x = n[s].v.substr(0, f);
					for (; f < n[s].v.length; ++f) {
						if (y < C.length) x += C.charAt(y++)
					}
					n[s].v = x;
					n[s].t = "t";
					A = s
				}
			}
		}
		for (s = 0; s < n.length; ++s) if (n[s] != null && "n?".indexOf(n[s].t) > -1) {
			S = a > 1 && r < 0 && s > 0 && n[s - 1].v === "-" ? -r: r;
			n[s].v = Ne(n[s].t, n[s].v, S);
			n[s].t = "t"
		}
		var T = "";
		for (s = 0; s !== n.length; ++s) if (n[s] != null) T += n[s].v;
		return T
	}
	var Be = /\[(=|>[=]?|<[>=]?)(-?\d+(?:\.\d*)?)\]/;
	function Ue(e, r) {
		if (r == null) return false;
		var t = parseFloat(r[2]);
		switch (r[1]) {
		case "=":
			if (e == t) return true;
			break;
		case ">":
			if (e > t) return true;
			break;
		case "<":
			if (e < t) return true;
			break;
		case "<>":
			if (e != t) return true;
			break;
		case ">=":
			if (e >= t) return true;
			break;
		case "<=":
			if (e <= t) return true;
			break
		}
		return false
	}
	function ze(e, r) {
		var t = Ie(e);
		var a = t.length,
		n = t[a - 1].indexOf("@");
		if (a < 4 && n > -1)--a;
		if (t.length > 4) throw new Error("cannot find right format for |" + t.join("|") + "|");
		if (typeof r !== "number") return [4, t.length === 4 || n > -1 ? t[t.length - 1] : "@"];
		switch (t.length) {
		case 1:
			t = n > -1 ? ["General", "General", "General", t[0]] : [t[0], t[0], t[0], "@"];
			break;
		case 2:
			t = n > -1 ? [t[0], t[0], t[0], t[1]] : [t[0], t[1], t[0], "@"];
			break;
		case 3:
			t = n > -1 ? [t[0], t[1], t[0], t[2]] : [t[0], t[1], t[2], "@"];
			break;
		case 4:
			break
		}
		var i = r > 0 ? t[0] : r < 0 ? t[1] : t[2];
		if (t[0].indexOf("[") === -1 && t[1].indexOf("[") === -1) return [a, i];
		if (t[0].match(/\[[=<>]/) != null || t[1].match(/\[[=<>]/) != null) {
			var s = t[0].match(Be);
			var l = t[1].match(Be);
			return Ue(r, s) ? [a, t[0]] : Ue(r, l) ? [a, t[1]] : [a, t[s != null && l != null ? 2 : 1]]
		}
		return [a, i]
	}
	function $e(e, r, t) {
		if (t == null) t = {};
		var a = "";
		switch (typeof e) {
		case "string":
			if (e == "m/d/yy" && t.dateNF) a = t.dateNF;
			else a = e;
			break;
		case "number":
			if (e == 14 && t.dateNF) a = t.dateNF;
			else a = (t.table != null ? t.table: q)[e];
			if (a == null) a = (t.table && t.table[Q[e]]) || q[Q[e]];
			if (a == null) a = ee[e] || "General";
			break
		}
		if (Y(a, 0)) return ce(r, t);
		if (r instanceof Date) r = dr(r, t.date1904);
		var n = ze(a, r);
		if (Y(n[1])) return ce(r, t);
		if (r === true) r = "TRUE";
		else if (r === false) r = "FALSE";
		else if (r === "" || r == null) return "";
		return Le(n[1], r, t, n[0])
	}
	function We(e, r) {
		if (typeof r != "number") {
			r = +r || -1;
			for (var t = 0; t < 392; ++t) {
				if (q[t] == undefined) {
					if (r < 0) r = t;
					continue
				}
				if (q[t] == e) {
					r = t;
					break
				}
			}
			if (r < 0) r = 391
		}
		q[r] = e;
		return r
	}
	function je(e) {
		for (var r = 0; r != 392; ++r) if (e[r] !== undefined) We(e[r], r)
	}
	function He() {
		q = K()
	}
	var Ve = {
		format: $e,
		load: We,
		_table: q,
		load_table: je,
		parse_date_code: ae,
		is_date: Re,
		get_table: function Rc() {
			return (Ve._table = q)
		},
	};
	var Xe = {
		5 : '"$"#,##0_);\\("$"#,##0\\)',
		6 : '"$"#,##0_);[Red]\\("$"#,##0\\)',
		7 : '"$"#,##0.00_);\\("$"#,##0.00\\)',
		8 : '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
		23 : "General",
		24 : "General",
		25 : "General",
		26 : "General",
		27 : "m/d/yy",
		28 : "m/d/yy",
		29 : "m/d/yy",
		30 : "m/d/yy",
		31 : "m/d/yy",
		32 : "h:mm:ss",
		33 : "h:mm:ss",
		34 : "h:mm:ss",
		35 : "h:mm:ss",
		36 : "m/d/yy",
		41 : '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
		42 : '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
		43 : '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
		44 : '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
		50 : "m/d/yy",
		51 : "m/d/yy",
		52 : "m/d/yy",
		53 : "m/d/yy",
		54 : "m/d/yy",
		55 : "m/d/yy",
		56 : "m/d/yy",
		57 : "m/d/yy",
		58 : "m/d/yy",
		59 : "0",
		60 : "0.00",
		61 : "#,##0",
		62 : "#,##0.00",
		63 : '"$"#,##0_);\\("$"#,##0\\)',
		64 : '"$"#,##0_);[Red]\\("$"#,##0\\)',
		65 : '"$"#,##0.00_);\\("$"#,##0.00\\)',
		66 : '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
		67 : "0%",
		68 : "0.00%",
		69 : "# ?/?",
		70 : "# ??/??",
		71 : "m/d/yy",
		72 : "m/d/yy",
		73 : "d-mmm-yy",
		74 : "d-mmm",
		75 : "mmm-yy",
		76 : "h:mm",
		77 : "h:mm:ss",
		78 : "m/d/yy h:mm",
		79 : "mm:ss",
		80 : "[h]:mm:ss",
		81 : "mmss.0",
	};
	var Ge = /[dD]+|[mM]+|[yYeE]+|[Hh]+|[Ss]+/g;
	function Ye(e) {
		var r = typeof e == "number" ? q[e] : e;
		r = r.replace(Ge, "(\\d+)");
		Ge.lastIndex = 0;
		return new RegExp("^" + r + "$")
	}
	function Je(e, r, t) {
		var a = -1,
		n = -1,
		i = -1,
		s = -1,
		l = -1,
		o = -1; (r.match(Ge) || []).forEach(function(e, r) {
			var c = parseInt(t[r + 1], 10);
			switch (e.toLowerCase().charAt(0)) {
			case "y":
				a = c;
				break;
			case "d":
				i = c;
				break;
			case "h":
				s = c;
				break;
			case "s":
				o = c;
				break;
			case "m":
				if (s >= 0) l = c;
				else n = c;
				break
			}
		});
		Ge.lastIndex = 0;
		if (o >= 0 && l == -1 && n >= 0) {
			l = n;
			n = -1
		}
		var c = ("" + (a >= 0 ? a: new Date().getFullYear())).slice(-4) + "-" + ("00" + (n >= 1 ? n: 1)).slice(-2) + "-" + ("00" + (i >= 1 ? i: 1)).slice(-2);
		if (c.length == 7) c = "0" + c;
		if (c.length == 8) c = "20" + c;
		var f = ("00" + (s >= 0 ? s: 0)).slice(-2) + ":" + ("00" + (l >= 0 ? l: 0)).slice(-2) + ":" + ("00" + (o >= 0 ? o: 0)).slice(-2);
		if (s == -1 && l == -1 && o == -1) return c;
		if (a == -1 && n == -1 && i == -1) return f;
		return c + "T" + f
	}
	var Ze = {
		"d.m": "d\\.m",
	};
	function Ke(e, r) {
		return We(Ze[e] || e, r)
	}
	var qe = (function() {
		var e = {};
		e.version = "1.2.0";
		function r() {
			var e = 0,
			r = new Array(256);
			for (var t = 0; t != 256; ++t) {
				e = t;
				e = e & 1 ? -306674912 ^ (e >>> 1) : e >>> 1;
				e = e & 1 ? -306674912 ^ (e >>> 1) : e >>> 1;
				e = e & 1 ? -306674912 ^ (e >>> 1) : e >>> 1;
				e = e & 1 ? -306674912 ^ (e >>> 1) : e >>> 1;
				e = e & 1 ? -306674912 ^ (e >>> 1) : e >>> 1;
				e = e & 1 ? -306674912 ^ (e >>> 1) : e >>> 1;
				e = e & 1 ? -306674912 ^ (e >>> 1) : e >>> 1;
				e = e & 1 ? -306674912 ^ (e >>> 1) : e >>> 1;
				r[t] = e
			}
			return typeof Int32Array !== "undefined" ? new Int32Array(r) : r
		}
		var t = r();
		function a(e) {
			var r = 0,
			t = 0,
			a = 0,
			n = typeof Int32Array !== "undefined" ? new Int32Array(4096) : new Array(4096);
			for (a = 0; a != 256; ++a) n[a] = e[a];
			for (a = 0; a != 256; ++a) {
				t = e[a];
				for (r = 256 + a; r < 4096; r += 256) t = n[r] = (t >>> 8) ^ e[t & 255]
			}
			var i = [];
			for (a = 1; a != 16; ++a) i[a - 1] = typeof Int32Array !== "undefined" && typeof n.subarray == "function" ? n.subarray(a * 256, a * 256 + 256) : n.slice(a * 256, a * 256 + 256);
			return i
		}
		var n = a(t);
		var i = n[0],
		s = n[1],
		l = n[2],
		o = n[3],
		c = n[4];
		var f = n[5],
		u = n[6],
		h = n[7],
		d = n[8],
		p = n[9];
		var m = n[10],
		v = n[11],
		g = n[12],
		b = n[13],
		w = n[14];
		function k(e, r) {
			var a = r ^ -1;
			for (var n = 0,
			i = e.length; n < i;) a = (a >>> 8) ^ t[(a ^ e.charCodeAt(n++)) & 255];
			return~a
		}
		function y(e, r) {
			var a = r ^ -1,
			n = e.length - 15,
			k = 0;
			for (; k < n;) a = w[e[k++] ^ (a & 255)] ^ b[e[k++] ^ ((a >> 8) & 255)] ^ g[e[k++] ^ ((a >> 16) & 255)] ^ v[e[k++] ^ (a >>> 24)] ^ m[e[k++]] ^ p[e[k++]] ^ d[e[k++]] ^ h[e[k++]] ^ u[e[k++]] ^ f[e[k++]] ^ c[e[k++]] ^ o[e[k++]] ^ l[e[k++]] ^ s[e[k++]] ^ i[e[k++]] ^ t[e[k++]];
			n += 15;
			while (k < n) a = (a >>> 8) ^ t[(a ^ e[k++]) & 255];
			return~a
		}
		function x(e, r) {
			var a = r ^ -1;
			for (var n = 0,
			i = e.length,
			s = 0,
			l = 0; n < i;) {
				s = e.charCodeAt(n++);
				if (s < 128) {
					a = (a >>> 8) ^ t[(a ^ s) & 255]
				} else if (s < 2048) {
					a = (a >>> 8) ^ t[(a ^ (192 | ((s >> 6) & 31))) & 255];
					a = (a >>> 8) ^ t[(a ^ (128 | (s & 63))) & 255]
				} else if (s >= 55296 && s < 57344) {
					s = (s & 1023) + 64;
					l = e.charCodeAt(n++) & 1023;
					a = (a >>> 8) ^ t[(a ^ (240 | ((s >> 8) & 7))) & 255];
					a = (a >>> 8) ^ t[(a ^ (128 | ((s >> 2) & 63))) & 255];
					a = (a >>> 8) ^ t[(a ^ (128 | ((l >> 6) & 15) | ((s & 3) << 4))) & 255];
					a = (a >>> 8) ^ t[(a ^ (128 | (l & 63))) & 255]
				} else {
					a = (a >>> 8) ^ t[(a ^ (224 | ((s >> 12) & 15))) & 255];
					a = (a >>> 8) ^ t[(a ^ (128 | ((s >> 6) & 63))) & 255];
					a = (a >>> 8) ^ t[(a ^ (128 | (s & 63))) & 255]
				}
			}
			return~a
		}
		e.table = t;
		e.bstr = k;
		e.buf = y;
		e.str = x;
		return e
	})();
	var Qe = (function Lc() {
		var e = {};
		e.version = "1.2.2";
		function r(e, r) {
			var t = e.split("/"),
			a = r.split("/");
			for (var n = 0,
			i = 0,
			s = Math.min(t.length, a.length); n < s; ++n) {
				if ((i = t[n].length - a[n].length)) return i;
				if (t[n] != a[n]) return t[n] < a[n] ? -1 : 1
			}
			return t.length - a.length
		}
		function t(e) {
			if (e.charAt(e.length - 1) == "/") return e.slice(0, -1).indexOf("/") === -1 ? e: t(e.slice(0, -1));
			var r = e.lastIndexOf("/");
			return r === -1 ? e: e.slice(0, r + 1)
		}
		function a(e) {
			if (e.charAt(e.length - 1) == "/") return a(e.slice(0, -1));
			var r = e.lastIndexOf("/");
			return r === -1 ? e: e.slice(r + 1)
		}
		function n(e, r) {
			if (typeof r === "string") r = new Date(r);
			var t = r.getHours();
			t = (t << 6) | r.getMinutes();
			t = (t << 5) | (r.getSeconds() >>> 1);
			e._W(2, t);
			var a = r.getFullYear() - 1980;
			a = (a << 4) | (r.getMonth() + 1);
			a = (a << 5) | r.getDate();
			e._W(2, a)
		}
		function i(e) {
			var r = e._R(2) & 65535;
			var t = e._R(2) & 65535;
			var a = new Date();
			var n = t & 31;
			t >>>= 5;
			var i = t & 15;
			t >>>= 4;
			a.setMilliseconds(0);
			a.setFullYear(t + 1980);
			a.setMonth(i - 1);
			a.setDate(n);
			var s = r & 31;
			r >>>= 5;
			var l = r & 63;
			r >>>= 6;
			a.setHours(r);
			a.setMinutes(l);
			a.setSeconds(s << 1);
			return a
		}
		function s(e) {
			ya(e, 0);
			var r = {};
			var t = 0;
			while (e.l <= e.length - 4) {
				var a = e._R(2);
				var n = e._R(2),
				i = e.l + n;
				var s = {};
				switch (a) {
				case 21589:
					{
						t = e._R(1);
						if (t & 1) s.mtime = e._R(4);
						if (n > 5) {
							if (t & 2) s.atime = e._R(4);
							if (t & 4) s.ctime = e._R(4)
						}
						if (s.mtime) s.mt = new Date(s.mtime * 1e3)
					}
					break;
				case 1:
					{
						var l = e._R(4),
						o = e._R(4);
						s.usz = o * Math.pow(2, 32) + l;
						l = e._R(4);
						o = e._R(4);
						s.csz = o * Math.pow(2, 32) + l
					}
					break
				}
				e.l = i;
				r[a] = s
			}
			return r
		}
		var l;
		function o() {
			return l || (l = er)
		}
		function c(e, r) {
			if (e[0] == 80 && e[1] == 75) return Oe(e, r);
			if ((e[0] | 32) == 109 && (e[1] | 32) == 105) return ze(e, r);
			if (e.length < 512) throw new Error("CFB file size " + e.length + " < 512");
			var t = 3;
			var a = 512;
			var n = 0;
			var i = 0;
			var s = 0;
			var l = 0;
			var o = 0;
			var c = [];
			var p = e.slice(0, 512);
			ya(p, 0);
			var v = f(p);
			t = v[0];
			switch (t) {
			case 3:
				a = 512;
				break;
			case 4:
				a = 4096;
				break;
			case 0:
				if (v[1] == 0) return Oe(e, r);
			default:
				throw new Error("Major Version: Expected 3 or 4 saw " + t);
			}
			if (a !== 512) {
				p = e.slice(0, a);
				ya(p, 28)
			}
			var w = e.slice(0, a);
			u(p, t);
			var k = p._R(4, "i");
			if (t === 3 && k !== 0) throw new Error("# Directory Sectors: Expected 0 saw " + k);
			p.l += 4;
			s = p._R(4, "i");
			p.l += 4;
			p.chk("00100000", "Mini Stream Cutoff Size: ");
			l = p._R(4, "i");
			n = p._R(4, "i");
			o = p._R(4, "i");
			i = p._R(4, "i");
			for (var y = -1,
			x = 0; x < 109; ++x) {
				y = p._R(4, "i");
				if (y < 0) break;
				c[x] = y
			}
			var S = h(e, a);
			m(o, i, S, a, c);
			var C = g(S, s, c, a);
			if (s < C.length) C[s].name = "!Directory";
			if (n > 0 && l !== R) C[l].name = "!MiniFAT";
			C[c[0]].name = "!FAT";
			C.fat_addrs = c;
			C.ssz = a;
			var _ = {},
			A = [],
			T = [],
			E = [];
			b(s, C, S, A, n, _, T, l);
			d(T, E, A);
			A.shift();
			var F = {
				FileIndex: T,
				FullPaths: E,
			};
			if (r && r.raw) F.raw = {
				header: w,
				sectors: S,
			};
			return F
		}
		function f(e) {
			if (e[e.l] == 80 && e[e.l + 1] == 75) return [0, 0];
			e.chk(U, "Header Signature: ");
			e.l += 16;
			var r = e._R(2, "u");
			return [e._R(2, "u"), r]
		}
		function u(e, r) {
			var t = 9;
			e.l += 2;
			switch ((t = e._R(2))) {
			case 9:
				if (r != 3) throw new Error("Sector Shift: Expected 9 saw " + t);
				break;
			case 12:
				if (r != 4) throw new Error("Sector Shift: Expected 12 saw " + t);
				break;
			default:
				throw new Error("Sector Shift: Expected 9 or 12 saw " + t);
			}
			e.chk("0600", "Mini Sector Shift: ");
			e.chk("000000000000", "Reserved: ")
		}
		function h(e, r) {
			var t = Math.ceil(e.length / r) - 1;
			var a = [];
			for (var n = 1; n < t; ++n) a[n - 1] = e.slice(n * r, (n + 1) * r);
			a[t - 1] = e.slice(t * r);
			return a
		}
		function d(e, r, t) {
			var a = 0,
			n = 0,
			i = 0,
			s = 0,
			l = 0,
			o = t.length;
			var c = [],
			f = [];
			for (; a < o; ++a) {
				c[a] = f[a] = a;
				r[a] = t[a]
			}
			for (; l < f.length; ++l) {
				a = f[l];
				n = e[a].L;
				i = e[a].R;
				s = e[a].C;
				if (c[a] === a) {
					if (n !== -1 && c[n] !== n) c[a] = c[n];
					if (i !== -1 && c[i] !== i) c[a] = c[i]
				}
				if (s !== -1) c[s] = a;
				if (n !== -1 && a != c[a]) {
					c[n] = c[a];
					if (f.lastIndexOf(n) < l) f.push(n)
				}
				if (i !== -1 && a != c[a]) {
					c[i] = c[a];
					if (f.lastIndexOf(i) < l) f.push(i)
				}
			}
			for (a = 1; a < o; ++a) if (c[a] === a) {
				if (i !== -1 && c[i] !== i) c[a] = c[i];
				else if (n !== -1 && c[n] !== n) c[a] = c[n]
			}
			for (a = 1; a < o; ++a) {
				if (e[a].type === 0) continue;
				l = a;
				if (l != c[l]) do {
					l = c[l];
					r[a] = r[l] + "/" + r[a]
				} while ( l !== 0 && - 1 !== c [ l ] && l != c[l]);
				c[a] = -1
			}
			r[0] += "/";
			for (a = 1; a < o; ++a) {
				if (e[a].type !== 2) r[a] += "/"
			}
		}
		function p(e, r, t) {
			var a = e.start,
			n = e.size;
			var i = [];
			var s = a;
			while (t && n > 0 && s >= 0) {
				i.push(r.slice(s * I, s * I + I));
				n -= I;
				s = da(t, s * 4)
			}
			if (i.length === 0) return Sa(0);
			return P(i).slice(0, e.size)
		}
		function m(e, r, t, a, n) {
			var i = R;
			if (e === R) {
				if (r !== 0) throw new Error("DIFAT chain shorter than expected");
			} else if (e !== -1) {
				var s = t[e],
				l = (a >>> 2) - 1;
				if (!s) return;
				for (var o = 0; o < l; ++o) {
					if ((i = da(s, o * 4)) === R) break;
					n.push(i)
				}
				if (r >= 1) m(da(s, a - 4), r - 1, t, a, n)
			}
		}
		function v(e, r, t, a, n) {
			var i = [],
			s = [];
			if (!n) n = [];
			var l = a - 1,
			o = 0,
			c = 0;
			for (o = r; o >= 0;) {
				n[o] = true;
				i[i.length] = o;
				s.push(e[o]);
				var f = t[Math.floor((o * 4) / a)];
				c = (o * 4) & l;
				if (a < 4 + c) throw new Error("FAT boundary crossed: " + o + " 4 " + a);
				if (!e[f]) break;
				o = da(e[f], c)
			}
			return {
				nodes: i,
				data: Wt([s]),
			}
		}
		function g(e, r, t, a) {
			var n = e.length,
			i = [];
			var s = [],
			l = [],
			o = [];
			var c = a - 1,
			f = 0,
			u = 0,
			h = 0,
			d = 0;
			for (f = 0; f < n; ++f) {
				l = [];
				h = f + r;
				if (h >= n) h -= n;
				if (s[h]) continue;
				o = [];
				var p = [];
				for (u = h; u >= 0;) {
					p[u] = true;
					s[u] = true;
					l[l.length] = u;
					o.push(e[u]);
					var m = t[Math.floor((u * 4) / a)];
					d = (u * 4) & c;
					if (a < 4 + d) throw new Error("FAT boundary crossed: " + u + " 4 " + a);
					if (!e[m]) break;
					u = da(e[m], d);
					if (p[u]) break
				}
				i[h] = {
					nodes: l,
					data: Wt([o]),
				}
			}
			return i
		}
		function b(e, r, t, a, n, i, s, l) {
			var o = 0,
			c = a.length ? 2 : 0;
			var f = r[e].data;
			var u = 0,
			h = 0,
			d;
			for (; u < f.length; u += 128) {
				var m = f.slice(u, u + 128);
				ya(m, 64);
				h = m._R(2);
				d = Ht(m, 0, h - c);
				a.push(d);
				var g = {
					name: d,
					type: m._R(1),
					color: m._R(1),
					L: m._R(4, "i"),
					R: m._R(4, "i"),
					C: m._R(4, "i"),
					clsid: m._R(16),
					state: m._R(4, "i"),
					start: 0,
					size: 0,
				};
				var b = m._R(2) + m._R(2) + m._R(2) + m._R(2);
				if (b !== 0) g.ct = w(m, m.l - 8);
				var k = m._R(2) + m._R(2) + m._R(2) + m._R(2);
				if (k !== 0) g.mt = w(m, m.l - 8);
				g.start = m._R(4, "i");
				g.size = m._R(4, "i");
				if (g.size < 0 && g.start < 0) {
					g.size = g.type = 0;
					g.start = R;
					g.name = ""
				}
				if (g.type === 5) {
					o = g.start;
					if (n > 0 && o !== R) r[o].name = "!StreamData"
				} else if (g.size >= 4096) {
					g.storage = "fat";
					if (r[g.start] === undefined) r[g.start] = v(t, g.start, r.fat_addrs, r.ssz);
					r[g.start].name = g.name;
					g.content = r[g.start].data.slice(0, g.size)
				} else {
					g.storage = "minifat";
					if (g.size < 0) g.size = 0;
					else if (o !== R && g.start !== R && r[o]) {
						g.content = p(g, r[o].data, (r[l] || {}).data)
					}
				}
				if (g.content) ya(g.content, 0);
				i[d] = g;
				s.push(g)
			}
		}
		function w(e, r) {
			return new Date(((ha(e, r + 4) / 1e7) * Math.pow(2, 32) + ha(e, r) / 1e7 - 11644473600) * 1e3)
		}
		function k(e, r) {
			o();
			return c(l.readFileSync(e), r)
		}
		function x(e, r) {
			var t = r && r.type;
			if (!t) {
				if (_ && Buffer.isBuffer(e)) t = "buffer"
			}
			switch (t || "base64") {
			case "file":
				return k(e, r);
			case "base64":
				return c(D(C(e)), r);
			case "binary":
				return c(D(e), r)
			}
			return c(e, r)
		}
		function S(e, r) {
			var t = r || {},
			a = t.root || "Root Entry";
			if (!e.FullPaths) e.FullPaths = [];
			if (!e.FileIndex) e.FileIndex = [];
			if (e.FullPaths.length !== e.FileIndex.length) throw new Error("inconsistent CFB structure");
			if (e.FullPaths.length === 0) {
				e.FullPaths[0] = a + "/";
				e.FileIndex[0] = {
					name: a,
					type: 5,
				}
			}
			if (t.CLSID) e.FileIndex[0].clsid = t.CLSID;
			T(e)
		}
		function T(e) {
			var r = "Sh33tJ5";
			if (Qe.find(e, "/" + r)) return;
			var t = Sa(4);
			t[0] = 55;
			t[1] = t[3] = 50;
			t[2] = 54;
			e.FileIndex.push({
				name: r,
				type: 2,
				content: t,
				size: 4,
				L: 69,
				R: 69,
				C: 69,
			});
			e.FullPaths.push(e.FullPaths[0] + r);
			O(e)
		}
		function O(e, n) {
			S(e);
			var i = false,
			s = false;
			for (var l = e.FullPaths.length - 1; l >= 0; --l) {
				var o = e.FileIndex[l];
				switch (o.type) {
				case 0:
					if (s) i = true;
					else {
						e.FileIndex.pop();
						e.FullPaths.pop()
					}
					break;
				case 1:
				case 2:
				case 5:
					s = true;
					if (isNaN(o.R * o.L * o.C)) i = true;
					if (o.R > -1 && o.L > -1 && o.R == o.L) i = true;
					break;
				default:
					i = true;
					break
				}
			}
			if (!i && !n) return;
			var c = new Date(1987, 1, 19),
			f = 0;
			var u = Object.create ? Object.create(null) : {};
			var h = [];
			for (l = 0; l < e.FullPaths.length; ++l) {
				u[e.FullPaths[l]] = true;
				if (e.FileIndex[l].type === 0) continue;
				h.push([e.FullPaths[l], e.FileIndex[l]])
			}
			for (l = 0; l < h.length; ++l) {
				var d = t(h[l][0]);
				s = u[d];
				while (!s) {
					while (t(d) && !u[t(d)]) d = t(d);
					h.push([d, {
						name: a(d).replace("/", ""),
						type: 1,
						clsid: $,
						ct: c,
						mt: c,
						content: null,
					},
					]);
					u[d] = true;
					d = t(h[l][0]);
					s = u[d]
				}
			}
			h.sort(function(e, t) {
				return r(e[0], t[0])
			});
			e.FullPaths = [];
			e.FileIndex = [];
			for (l = 0; l < h.length; ++l) {
				e.FullPaths[l] = h[l][0];
				e.FileIndex[l] = h[l][1]
			}
			for (l = 0; l < h.length; ++l) {
				var p = e.FileIndex[l];
				var m = e.FullPaths[l];
				p.name = a(m).replace("/", "");
				p.L = p.R = p.C = -(p.color = 1);
				p.size = p.content ? p.content.length: 0;
				p.start = 0;
				p.clsid = p.clsid || $;
				if (l === 0) {
					p.C = h.length > 1 ? 1 : -1;
					p.size = 0;
					p.type = 5
				} else if (m.slice(-1) == "/") {
					for (f = l + 1; f < h.length; ++f) if (t(e.FullPaths[f]) == m) break;
					p.C = f >= h.length ? -1 : f;
					for (f = l + 1; f < h.length; ++f) if (t(e.FullPaths[f]) == t(m)) break;
					p.R = f >= h.length ? -1 : f;
					p.type = 1
				} else {
					if (t(e.FullPaths[l + 1] || "") == t(m)) p.R = l + 1;
					p.type = 2
				}
			}
		}
		function M(e, r) {
			var t = r || {};
			if (t.fileType == "mad") return $e(e, t);
			O(e);
			switch (t.fileType) {
			case "zip":
				return Ne(e, t)
			}
			var a = (function(e) {
				var r = 0,
				t = 0;
				for (var a = 0; a < e.FileIndex.length; ++a) {
					var n = e.FileIndex[a];
					if (!n.content) continue;
					var i = n.content.length;
					if (i > 0) {
						if (i < 4096) r += (i + 63) >> 6;
						else t += (i + 511) >> 9
					}
				}
				var s = (e.FullPaths.length + 3) >> 2;
				var l = (r + 7) >> 3;
				var o = (r + 127) >> 7;
				var c = l + t + s + o;
				var f = (c + 127) >> 7;
				var u = f <= 109 ? 0 : Math.ceil((f - 109) / 127);
				while ((c + f + u + 127) >> 7 > f) u = ++f <= 109 ? 0 : Math.ceil((f - 109) / 127);
				var h = [1, u, f, o, s, t, r, 0];
				e.FileIndex[0].size = r << 6;
				h[7] = (e.FileIndex[0].start = h[0] + h[1] + h[2] + h[3] + h[4] + h[5]) + ((h[6] + 7) >> 3);
				return h
			})(e);
			var n = Sa(a[7] << 9);
			var i = 0,
			s = 0; {
				for (i = 0; i < 8; ++i) n._W(1, z[i]);
				for (i = 0; i < 8; ++i) n._W(2, 0);
				n._W(2, 62);
				n._W(2, 3);
				n._W(2, 65534);
				n._W(2, 9);
				n._W(2, 6);
				for (i = 0; i < 3; ++i) n._W(2, 0);
				n._W(4, 0);
				n._W(4, a[2]);
				n._W(4, a[0] + a[1] + a[2] + a[3] - 1);
				n._W(4, 0);
				n._W(4, 1 << 12);
				n._W(4, a[3] ? a[0] + a[1] + a[2] - 1 : R);
				n._W(4, a[3]);
				n._W(-4, a[1] ? a[0] - 1 : R);
				n._W(4, a[1]);
				for (i = 0; i < 109; ++i) n._W(-4, i < a[2] ? a[1] + i: -1)
			}
			if (a[1]) {
				for (s = 0; s < a[1]; ++s) {
					for (; i < 236 + s * 127; ++i) n._W(-4, i < a[2] ? a[1] + i: -1);
					n._W(-4, s === a[1] - 1 ? R: s + 1)
				}
			}
			var l = function(e) {
				for (s += e; i < s - 1; ++i) n._W(-4, i + 1);
				if (e) {++i;
					n._W(-4, R)
				}
			};
			s = i = 0;
			for (s += a[1]; i < s; ++i) n._W(-4, W.DIFSECT);
			for (s += a[2]; i < s; ++i) n._W(-4, W.FATSECT);
			l(a[3]);
			l(a[4]);
			var o = 0,
			c = 0;
			var f = e.FileIndex[0];
			for (; o < e.FileIndex.length; ++o) {
				f = e.FileIndex[o];
				if (!f.content) continue;
				c = f.content.length;
				if (c < 4096) continue;
				f.start = s;
				l((c + 511) >> 9)
			}
			l((a[6] + 7) >> 3);
			while (n.l & 511) n._W(-4, W.ENDOFCHAIN);
			s = i = 0;
			for (o = 0; o < e.FileIndex.length; ++o) {
				f = e.FileIndex[o];
				if (!f.content) continue;
				c = f.content.length;
				if (!c || c >= 4096) continue;
				f.start = s;
				l((c + 63) >> 6)
			}
			while (n.l & 511) n._W(-4, W.ENDOFCHAIN);
			for (i = 0; i < a[4] << 2; ++i) {
				var u = e.FullPaths[i];
				if (!u || u.length === 0) {
					for (o = 0; o < 17; ++o) n._W(4, 0);
					for (o = 0; o < 3; ++o) n._W(4, -1);
					for (o = 0; o < 12; ++o) n._W(4, 0);
					continue
				}
				f = e.FileIndex[i];
				if (i === 0) f.start = f.size ? f.start - 1 : R;
				var h = (i === 0 && t.root) || f.name;
				if (h.length > 32) {
					console.error("Name " + h + " will be truncated to " + h.slice(0, 32));
					h = h.slice(0, 32)
				}
				c = 2 * (h.length + 1);
				n._W(64, h, "utf16le");
				n._W(2, c);
				n._W(1, f.type);
				n._W(1, f.color);
				n._W(-4, f.L);
				n._W(-4, f.R);
				n._W(-4, f.C);
				if (!f.clsid) for (o = 0; o < 4; ++o) n._W(4, 0);
				else n._W(16, f.clsid, "hex");
				n._W(4, f.state || 0);
				n._W(4, 0);
				n._W(4, 0);
				n._W(4, 0);
				n._W(4, 0);
				n._W(4, f.start);
				n._W(4, f.size);
				n._W(4, 0)
			}
			for (i = 1; i < e.FileIndex.length; ++i) {
				f = e.FileIndex[i];
				if (f.size >= 4096) {
					n.l = (f.start + 1) << 9;
					if (_ && Buffer.isBuffer(f.content)) {
						f.content.copy(n, n.l, 0, f.size);
						n.l += (f.size + 511) & -512
					} else {
						for (o = 0; o < f.size; ++o) n._W(1, f.content[o]);
						for (; o & 511; ++o) n._W(1, 0)
					}
				}
			}
			for (i = 1; i < e.FileIndex.length; ++i) {
				f = e.FileIndex[i];
				if (f.size > 0 && f.size < 4096) {
					if (_ && Buffer.isBuffer(f.content)) {
						f.content.copy(n, n.l, 0, f.size);
						n.l += (f.size + 63) & -64
					} else {
						for (o = 0; o < f.size; ++o) n._W(1, f.content[o]);
						for (; o & 63; ++o) n._W(1, 0)
					}
				}
			}
			if (_) {
				n.l = n.length
			} else {
				while (n.l < n.length) n._W(1, 0)
			}
			return n
		}
		function N(e, r) {
			var t = e.FullPaths.map(function(e) {
				return e.toUpperCase()
			});
			var a = t.map(function(e) {
				var r = e.split("/");
				return r[r.length - (e.slice(-1) == "/" ? 2 : 1)]
			});
			var n = false;
			if (r.charCodeAt(0) === 47) {
				n = true;
				r = t[0].slice(0, -1) + r
			} else n = r.indexOf("/") !== -1;
			var i = r.toUpperCase();
			var s = n === true ? t.indexOf(i) : a.indexOf(i);
			if (s !== -1) return e.FileIndex[s];
			var l = !i.match(B);
			i = i.replace(L, "");
			if (l) i = i.replace(B, "!");
			for (s = 0; s < t.length; ++s) {
				if ((l ? t[s].replace(B, "!") : t[s]).replace(L, "") == i) return e.FileIndex[s];
				if ((l ? a[s].replace(B, "!") : a[s]).replace(L, "") == i) return e.FileIndex[s]
			}
			return null
		}
		var I = 64;
		var R = -2;
		var U = "d0cf11e0a1b11ae1";
		var z = [208, 207, 17, 224, 161, 177, 26, 225];
		var $ = "00000000000000000000000000000000";
		var W = {
			MAXREGSECT: -6,
			DIFSECT: -4,
			FATSECT: -3,
			ENDOFCHAIN: R,
			FREESECT: -1,
			HEADER_SIGNATURE: U,
			HEADER_MINOR_VERSION: "3e00",
			MAXREGSID: -6,
			NOSTREAM: -1,
			HEADER_CLSID: $,
			EntryTypes: ["unknown", "storage", "stream", "lockbytes", "property", "root", ],
		};
		function j(e, r, t) {
			o();
			var a = M(e, t);
			l.writeFileSync(r, a)
		}
		function H(e) {
			var r = new Array(e.length);
			for (var t = 0; t < e.length; ++t) r[t] = String.fromCharCode(e[t]);
			return r.join("")
		}
		function V(e, r) {
			var t = M(e, r);
			switch ((r && r.type) || "buffer") {
			case "file":
				o();
				l.writeFileSync(r.filename, t);
				return t;
			case "binary":
				return typeof t == "string" ? t: H(t);
			case "base64":
				return y(typeof t == "string" ? t: H(t));
			case "buffer":
				if (_) return Buffer.isBuffer(t) ? t: A(t);
			case "array":
				return typeof t == "string" ? D(t) : t
			}
			return t
		}
		var X;
		function G(e) {
			try {
				var r = e.InflateRaw;
				var t = new r();
				t._processChunk(new Uint8Array([3, 0]), t._finishFlushFlag);
				if (t.bytesRead) X = e;
				else throw new Error("zlib does not expose bytesRead");
			} catch(a) {
				console.error("cannot use native zlib: " + (a.message || a))
			}
		}
		function Y(e, r) {
			if (!X) return Fe(e, r);
			var t = X.InflateRaw;
			var a = new t();
			var n = a._processChunk(e.slice(e.l), a._finishFlushFlag);
			e.l += a.bytesRead;
			return n
		}
		function J(e) {
			return X ? X.deflateRawSync(e) : ye(e)
		}
		var Z = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
		var K = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258, ];
		var q = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, ];
		function Q(e) {
			var r = (((e << 1) | (e << 11)) & 139536) | (((e << 5) | (e << 15)) & 558144);
			return ((r >> 16) | (r >> 8) | r) & 255
		}
		var ee = typeof Uint8Array !== "undefined";
		var re = ee ? new Uint8Array(1 << 8) : [];
		for (var te = 0; te < 1 << 8; ++te) re[te] = Q(te);
		function ae(e, r) {
			var t = re[e & 255];
			if (r <= 8) return t >>> (8 - r);
			t = (t << 8) | re[(e >> 8) & 255];
			if (r <= 16) return t >>> (16 - r);
			t = (t << 8) | re[(e >> 16) & 255];
			return t >>> (24 - r)
		}
		function ne(e, r) {
			var t = r & 7,
			a = r >>> 3;
			return ((e[a] | (t <= 6 ? 0 : e[a + 1] << 8)) >>> t) & 3
		}
		function ie(e, r) {
			var t = r & 7,
			a = r >>> 3;
			return ((e[a] | (t <= 5 ? 0 : e[a + 1] << 8)) >>> t) & 7
		}
		function se(e, r) {
			var t = r & 7,
			a = r >>> 3;
			return ((e[a] | (t <= 4 ? 0 : e[a + 1] << 8)) >>> t) & 15
		}
		function le(e, r) {
			var t = r & 7,
			a = r >>> 3;
			return ((e[a] | (t <= 3 ? 0 : e[a + 1] << 8)) >>> t) & 31
		}
		function oe(e, r) {
			var t = r & 7,
			a = r >>> 3;
			return ((e[a] | (t <= 1 ? 0 : e[a + 1] << 8)) >>> t) & 127
		}
		function ce(e, r, t) {
			var a = r & 7,
			n = r >>> 3,
			i = (1 << t) - 1;
			var s = e[n] >>> a;
			if (t < 8 - a) return s & i;
			s |= e[n + 1] << (8 - a);
			if (t < 16 - a) return s & i;
			s |= e[n + 2] << (16 - a);
			if (t < 24 - a) return s & i;
			s |= e[n + 3] << (24 - a);
			return s & i
		}
		function fe(e, r, t) {
			var a = r & 7,
			n = r >>> 3;
			if (a <= 5) e[n] |= (t & 7) << a;
			else {
				e[n] |= (t << a) & 255;
				e[n + 1] = (t & 7) >> (8 - a)
			}
			return r + 3
		}
		function ue(e, r, t) {
			var a = r & 7,
			n = r >>> 3;
			t = (t & 1) << a;
			e[n] |= t;
			return r + 1
		}
		function he(e, r, t) {
			var a = r & 7,
			n = r >>> 3;
			t <<= a;
			e[n] |= t & 255;
			t >>>= 8;
			e[n + 1] = t;
			return r + 8
		}
		function de(e, r, t) {
			var a = r & 7,
			n = r >>> 3;
			t <<= a;
			e[n] |= t & 255;
			t >>>= 8;
			e[n + 1] = t & 255;
			e[n + 2] = t >>> 8;
			return r + 16
		}
		function pe(e, r) {
			var t = e.length,
			a = 2 * t > r ? 2 * t: r + 5,
			n = 0;
			if (t >= r) return e;
			if (_) {
				var i = F(a);
				if (e.copy) e.copy(i);
				else for (; n < e.length; ++n) i[n] = e[n];
				return i
			} else if (ee) {
				var s = new Uint8Array(a);
				if (s.set) s.set(e);
				else for (; n < t; ++n) s[n] = e[n];
				return s
			}
			e.length = a;
			return e
		}
		function me(e) {
			var r = new Array(e);
			for (var t = 0; t < e; ++t) r[t] = 0;
			return r
		}
		function ve(e, r, t) {
			var a = 1,
			n = 0,
			i = 0,
			s = 0,
			l = 0,
			o = e.length;
			var c = ee ? new Uint16Array(32) : me(32);
			for (i = 0; i < 32; ++i) c[i] = 0;
			for (i = o; i < t; ++i) e[i] = 0;
			o = e.length;
			var f = ee ? new Uint16Array(o) : me(o);
			for (i = 0; i < o; ++i) {
				c[(n = e[i])]++;
				if (a < n) a = n;
				f[i] = 0
			}
			c[0] = 0;
			for (i = 1; i <= a; ++i) c[i + 16] = l = (l + c[i - 1]) << 1;
			for (i = 0; i < o; ++i) {
				l = e[i];
				if (l != 0) f[i] = c[l + 16]++
			}
			var u = 0;
			for (i = 0; i < o; ++i) {
				u = e[i];
				if (u != 0) {
					l = ae(f[i], a) >> (a - u);
					for (s = (1 << (a + 4 - u)) - 1; s >= 0; --s) r[l | (s << u)] = (u & 15) | (i << 4)
				}
			}
			return a
		}
		var ge = ee ? new Uint16Array(512) : me(512);
		var be = ee ? new Uint16Array(32) : me(32);
		if (!ee) {
			for (var we = 0; we < 512; ++we) ge[we] = 0;
			for (we = 0; we < 32; ++we) be[we] = 0
		} (function() {
			var e = [];
			var r = 0;
			for (; r < 32; r++) e.push(5);
			ve(e, be, 32);
			var t = [];
			r = 0;
			for (; r <= 143; r++) t.push(8);
			for (; r <= 255; r++) t.push(9);
			for (; r <= 279; r++) t.push(7);
			for (; r <= 287; r++) t.push(8);
			ve(t, ge, 288)
		})();
		var ke = (function Ge() {
			var e = ee ? new Uint8Array(32768) : [];
			var r = 0,
			t = 0;
			for (; r < q.length - 1; ++r) {
				for (; t < q[r + 1]; ++t) e[t] = r
			}
			for (; t < 32768; ++t) e[t] = 29;
			var a = ee ? new Uint8Array(259) : [];
			for (r = 0, t = 0; r < K.length - 1; ++r) {
				for (; t < K[r + 1]; ++t) a[t] = r
			}
			function n(e, r) {
				var t = 0;
				while (t < e.length) {
					var a = Math.min(65535, e.length - t);
					var n = t + a == e.length;
					r._W(1, +n);
					r._W(2, a);
					r._W(2, ~a & 65535);
					while (a-->0) r[r.l++] = e[t++]
				}
				return r.l
			}
			function i(r, t) {
				var n = 0;
				var i = 0;
				var s = ee ? new Uint16Array(32768) : [];
				while (i < r.length) {
					var l = Math.min(65535, r.length - i);
					if (l < 10) {
						n = fe(t, n, +!!(i + l == r.length));
						if (n & 7) n += 8 - (n & 7);
						t.l = (n / 8) | 0;
						t._W(2, l);
						t._W(2, ~l & 65535);
						while (l-->0) t[t.l++] = r[i++];
						n = t.l * 8;
						continue
					}
					n = fe(t, n, +!!(i + l == r.length) + 2);
					var o = 0;
					while (l-->0) {
						var c = r[i];
						o = ((o << 5) ^ c) & 32767;
						var f = -1,
						u = 0;
						if ((f = s[o])) {
							f |= i & ~32767;
							if (f > i) f -= 32768;
							if (f < i) while (r[f + u] == r[i + u] && u < 250)++u
						}
						if (u > 2) {
							c = a[u];
							if (c <= 22) n = he(t, n, re[c + 1] >> 1) - 1;
							else {
								he(t, n, 3);
								n += 5;
								he(t, n, re[c - 23] >> 5);
								n += 3
							}
							var h = c < 8 ? 0 : (c - 4) >> 2;
							if (h > 0) {
								de(t, n, u - K[c]);
								n += h
							}
							c = e[i - f];
							n = he(t, n, re[c] >> 3);
							n -= 3;
							var d = c < 4 ? 0 : (c - 2) >> 1;
							if (d > 0) {
								de(t, n, i - f - q[c]);
								n += d
							}
							for (var p = 0; p < u; ++p) {
								s[o] = i & 32767;
								o = ((o << 5) ^ r[i]) & 32767; ++i
							}
							l -= u - 1
						} else {
							if (c <= 143) c = c + 48;
							else n = ue(t, n, 1);
							n = he(t, n, re[c]);
							s[o] = i & 32767; ++i
						}
					}
					n = he(t, n, 0) - 1
				}
				t.l = ((n + 7) / 8) | 0;
				return t.l
			}
			return function s(e, r) {
				if (e.length < 8) return n(e, r);
				return i(e, r)
			}
		})();
		function ye(e) {
			var r = Sa(50 + Math.floor(e.length * 1.1));
			var t = ke(e, r);
			return r.slice(0, t)
		}
		var xe = ee ? new Uint16Array(32768) : me(32768);
		var Se = ee ? new Uint16Array(32768) : me(32768);
		var Ce = ee ? new Uint16Array(128) : me(128);
		var _e = 1,
		Ae = 1;
		function Te(e, r) {
			var t = le(e, r) + 257;
			r += 5;
			var a = le(e, r) + 1;
			r += 5;
			var n = se(e, r) + 4;
			r += 4;
			var i = 0;
			var s = ee ? new Uint8Array(19) : me(19);
			var l = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
			var o = 1;
			var c = ee ? new Uint8Array(8) : me(8);
			var f = ee ? new Uint8Array(8) : me(8);
			var u = s.length;
			for (var h = 0; h < n; ++h) {
				s[Z[h]] = i = ie(e, r);
				if (o < i) o = i;
				c[i]++;
				r += 3
			}
			var d = 0;
			c[0] = 0;
			for (h = 1; h <= o; ++h) f[h] = d = (d + c[h - 1]) << 1;
			for (h = 0; h < u; ++h) if ((d = s[h]) != 0) l[h] = f[d]++;
			var p = 0;
			for (h = 0; h < u; ++h) {
				p = s[h];
				if (p != 0) {
					d = re[l[h]] >> (8 - p);
					for (var m = (1 << (7 - p)) - 1; m >= 0; --m) Ce[d | (m << p)] = (p & 7) | (h << 3)
				}
			}
			var v = [];
			o = 1;
			for (; v.length < t + a;) {
				d = Ce[oe(e, r)];
				r += d & 7;
				switch ((d >>>= 3)) {
				case 16:
					i = 3 + ne(e, r);
					r += 2;
					d = v[v.length - 1];
					while (i-->0) v.push(d);
					break;
				case 17:
					i = 3 + ie(e, r);
					r += 3;
					while (i-->0) v.push(0);
					break;
				case 18:
					i = 11 + oe(e, r);
					r += 7;
					while (i-->0) v.push(0);
					break;
				default:
					v.push(d);
					if (o < d) o = d;
					break
				}
			}
			var g = v.slice(0, t),
			b = v.slice(t);
			for (h = t; h < 286; ++h) g[h] = 0;
			for (h = a; h < 30; ++h) b[h] = 0;
			_e = ve(g, xe, 286);
			Ae = ve(b, Se, 30);
			return r
		}
		function Ee(e, r) {
			if (e[0] == 3 && !(e[1] & 3)) {
				return [E(r), 2]
			}
			var t = 0;
			var a = 0;
			var n = F(r ? r: 1 << 18);
			var i = 0;
			var s = n.length >>> 0;
			var l = 0,
			o = 0;
			while ((a & 1) == 0) {
				a = ie(e, t);
				t += 3;
				if (a >>> 1 == 0) {
					if (t & 7) t += 8 - (t & 7);
					var c = e[t >>> 3] | (e[(t >>> 3) + 1] << 8);
					t += 32;
					if (c > 0) {
						if (!r && s < i + c) {
							n = pe(n, i + c);
							s = n.length
						}
						while (c-->0) {
							n[i++] = e[t >>> 3];
							t += 8
						}
					}
					continue
				} else if (a >> 1 == 1) {
					l = 9;
					o = 5
				} else {
					t = Te(e, t);
					l = _e;
					o = Ae
				}
				for (;;) {
					if (!r && s < i + 32767) {
						n = pe(n, i + 32767);
						s = n.length
					}
					var f = ce(e, t, l);
					var u = a >>> 1 == 1 ? ge[f] : xe[f];
					t += u & 15;
					u >>>= 4;
					if (((u >>> 8) & 255) === 0) n[i++] = u;
					else if (u == 256) break;
					else {
						u -= 257;
						var h = u < 8 ? 0 : (u - 4) >> 2;
						if (h > 5) h = 0;
						var d = i + K[u];
						if (h > 0) {
							d += ce(e, t, h);
							t += h
						}
						f = ce(e, t, o);
						u = a >>> 1 == 1 ? be[f] : Se[f];
						t += u & 15;
						u >>>= 4;
						var p = u < 4 ? 0 : (u - 2) >> 1;
						var m = q[u];
						if (p > 0) {
							m += ce(e, t, p);
							t += p
						}
						if (!r && s < d) {
							n = pe(n, d + 100);
							s = n.length
						}
						while (i < d) {
							n[i] = n[i - m]; ++i
						}
					}
				}
			}
			if (r) return [n, (t + 7) >>> 3];
			return [n.slice(0, i), (t + 7) >>> 3]
		}
		function Fe(e, r) {
			var t = e.slice(e.l || 0);
			var a = Ee(t, r);
			e.l += a[1];
			return a[0]
		}
		function De(e, r) {
			if (e) {
				if (typeof console !== "undefined") console.error(r)
			} else throw new Error(r);
		}
		function Oe(e, r) {
			var t = e;
			ya(t, 0);
			var a = [],
			n = [];
			var i = {
				FileIndex: a,
				FullPaths: n,
			};
			S(i, {
				root: r.root,
			});
			var l = t.length - 4;
			while ((t[l] != 80 || t[l + 1] != 75 || t[l + 2] != 5 || t[l + 3] != 6) && l >= 0)--l;
			t.l = l + 4;
			t.l += 4;
			var o = t._R(2);
			t.l += 6;
			var c = t._R(4);
			t.l = c;
			for (l = 0; l < o; ++l) {
				t.l += 20;
				var f = t._R(4);
				var u = t._R(4);
				var h = t._R(2);
				var d = t._R(2);
				var p = t._R(2);
				t.l += 8;
				var m = t._R(4);
				var v = s(t.slice(t.l + h, t.l + h + d));
				t.l += h + d + p;
				var g = t.l;
				t.l = m + 4;
				if (v && v[1]) {
					if ((v[1] || {}).usz) u = v[1].usz;
					if ((v[1] || {}).csz) f = v[1].csz
				}
				Me(t, f, u, i, v);
				t.l = g
			}
			return i
		}
		function Me(e, r, t, a, n) {
			e.l += 2;
			var l = e._R(2);
			var o = e._R(2);
			var c = i(e);
			if (l & 8257) throw new Error("Unsupported ZIP encryption");
			var f = e._R(4);
			var u = e._R(4);
			var h = e._R(4);
			var d = e._R(2);
			var p = e._R(2);
			var m = "";
			for (var v = 0; v < d; ++v) m += String.fromCharCode(e[e.l++]);
			if (p) {
				var g = s(e.slice(e.l, e.l + p));
				if ((g[21589] || {}).mt) c = g[21589].mt;
				if ((g[1] || {}).usz) h = g[1].usz;
				if ((g[1] || {}).csz) u = g[1].csz;
				if (n) {
					if ((n[21589] || {}).mt) c = n[21589].mt;
					if ((n[1] || {}).usz) h = g[1].usz;
					if ((n[1] || {}).csz) u = g[1].csz
				}
			}
			e.l += p;
			var b = e.slice(e.l, e.l + u);
			switch (o) {
			case 8:
				b = Y(e, h);
				break;
			case 0:
				break;
			default:
				throw new Error("Unsupported ZIP Compression method " + o);
			}
			var w = false;
			if (l & 8) {
				f = e._R(4);
				if (f == 134695760) {
					f = e._R(4);
					w = true
				}
				u = e._R(4);
				h = e._R(4)
			}
			if (u != r) De(w, "Bad compressed size: " + r + " != " + u);
			if (h != t) De(w, "Bad uncompressed size: " + t + " != " + h);
			je(a, m, b, {
				unsafe: true,
				mt: c,
			})
		}
		function Ne(e, r) {
			var t = r || {};
			var a = [],
			i = [];
			var s = Sa(1);
			var l = t.compression ? 8 : 0,
			o = 0;
			var c = false;
			if (c) o |= 8;
			var f = 0,
			u = 0;
			var h = 0,
			d = 0;
			var p = e.FullPaths[0],
			m = p,
			v = e.FileIndex[0];
			var g = [];
			var b = 0;
			for (f = 1; f < e.FullPaths.length; ++f) {
				m = e.FullPaths[f].slice(p.length);
				v = e.FileIndex[f];
				if (!v.size || !v.content || m == "Sh33tJ5") continue;
				var w = h;
				var k = Sa(m.length);
				for (u = 0; u < m.length; ++u) k._W(1, m.charCodeAt(u) & 127);
				k = k.slice(0, k.l);
				g[d] = typeof v.content == "string" ? qe.bstr(v.content, 0) : qe.buf(v.content, 0);
				var y = typeof v.content == "string" ? D(v.content) : v.content;
				if (l == 8) y = J(y);
				s = Sa(30);
				s._W(4, 67324752);
				s._W(2, 20);
				s._W(2, o);
				s._W(2, l);
				if (v.mt) n(s, v.mt);
				else s._W(4, 0);
				s._W(-4, o & 8 ? 0 : g[d]);
				s._W(4, o & 8 ? 0 : y.length);
				s._W(4, o & 8 ? 0 : v.content.length);
				s._W(2, k.length);
				s._W(2, 0);
				h += s.length;
				a.push(s);
				h += k.length;
				a.push(k);
				h += y.length;
				a.push(y);
				if (o & 8) {
					s = Sa(12);
					s._W(-4, g[d]);
					s._W(4, y.length);
					s._W(4, v.content.length);
					h += s.l;
					a.push(s)
				}
				s = Sa(46);
				s._W(4, 33639248);
				s._W(2, 0);
				s._W(2, 20);
				s._W(2, o);
				s._W(2, l);
				s._W(4, 0);
				s._W(-4, g[d]);
				s._W(4, y.length);
				s._W(4, v.content.length);
				s._W(2, k.length);
				s._W(2, 0);
				s._W(2, 0);
				s._W(2, 0);
				s._W(2, 0);
				s._W(4, 0);
				s._W(4, w);
				b += s.l;
				i.push(s);
				b += k.length;
				i.push(k); ++d
			}
			s = Sa(22);
			s._W(4, 101010256);
			s._W(2, 0);
			s._W(2, 0);
			s._W(2, d);
			s._W(2, d);
			s._W(4, b);
			s._W(4, h);
			s._W(2, 0);
			return P([P(a), P(i), s])
		}
		var Ie = {
			htm: "text/html",
			xml: "text/xml",
			gif: "image/gif",
			jpg: "image/jpeg",
			png: "image/png",
			mso: "application/x-mso",
			thmx: "application/vnd.ms-officetheme",
			sh33tj5: "application/octet-stream",
		};
		function Pe(e, r) {
			if (e.ctype) return e.ctype;
			var t = e.name || "",
			a = t.match(/\.([^\.]+)$/);
			if (a && Ie[a[1]]) return Ie[a[1]];
			if (r) {
				a = (t = r).match(/[\.\\]([^\.\\])+$/);
				if (a && Ie[a[1]]) return Ie[a[1]]
			}
			return "application/octet-stream"
		}
		function Re(e) {
			var r = y(e);
			var t = [];
			for (var a = 0; a < r.length; a += 76) t.push(r.slice(a, a + 76));
			return t.join("\r\n") + "\r\n"
		}
		function Le(e) {
			var r = e.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7E-\xFF=]/g,
			function(e) {
				var r = e.charCodeAt(0).toString(16).toUpperCase();
				return "=" + (r.length == 1 ? "0" + r: r)
			});
			r = r.replace(/ $/gm, "=20").replace(/\t$/gm, "=09");
			if (r.charAt(0) == "\n") r = "=0D" + r.slice(1);
			r = r.replace(/\r(?!\n)/gm, "=0D").replace(/\n\n/gm, "\n=0A").replace(/([^\r\n])\n/gm, "$1=0A");
			var t = [],
			a = r.split("\r\n");
			for (var n = 0; n < a.length; ++n) {
				var i = a[n];
				if (i.length == 0) {
					t.push("");
					continue
				}
				for (var s = 0; s < i.length;) {
					var l = 76;
					var o = i.slice(s, s + l);
					if (o.charAt(l - 1) == "=") l--;
					else if (o.charAt(l - 2) == "=") l -= 2;
					else if (o.charAt(l - 3) == "=") l -= 3;
					o = i.slice(s, s + l);
					s += l;
					if (s < i.length) o += "=";
					t.push(o)
				}
			}
			return t.join("\r\n")
		}
		function Be(e) {
			var r = [];
			for (var t = 0; t < e.length; ++t) {
				var a = e[t];
				while (t <= e.length && a.charAt(a.length - 1) == "=") a = a.slice(0, a.length - 1) + e[++t];
				r.push(a)
			}
			for (var n = 0; n < r.length; ++n) r[n] = r[n].replace(/[=][0-9A-Fa-f]{2}/g,
			function(e) {
				return String.fromCharCode(parseInt(e.slice(1), 16))
			});
			return D(r.join("\r\n"))
		}
		function Ue(e, r, t) {
			var a = "",
			n = "",
			i = "",
			s;
			var l = 0;
			for (; l < 10; ++l) {
				var o = r[l];
				if (!o || o.match(/^\s*$/)) break;
				var c = o.match(/^(.*?):\s*([^\s].*)$/);
				if (c) switch (c[1].toLowerCase()) {
				case "content-location":
					a = c[2].trim();
					break;
				case "content-type":
					i = c[2].trim();
					break;
				case "content-transfer-encoding":
					n = c[2].trim();
					break
				}
			}++l;
			switch (n.toLowerCase()) {
			case "base64":
				s = D(C(r.slice(l).join("")));
				break;
			case "quoted-printable":
				s = Be(r.slice(l));
				break;
			default:
				throw new Error("Unsupported Content-Transfer-Encoding " + n);
			}
			var f = je(e, a.slice(t.length), s, {
				unsafe: true,
			});
			if (i) f.ctype = i
		}
		function ze(e, r) {
			if (H(e.slice(0, 13)).toLowerCase() != "mime-version:") throw new Error("Unsupported MAD header");
			var t = (r && r.root) || "";
			var a = (_ && Buffer.isBuffer(e) ? e.toString("binary") : H(e)).split("\r\n");
			var n = 0,
			i = "";
			for (n = 0; n < a.length; ++n) {
				i = a[n];
				if (!/^Content-Location:/i.test(i)) continue;
				i = i.slice(i.indexOf("file"));
				if (!t) t = i.slice(0, i.lastIndexOf("/") + 1);
				if (i.slice(0, t.length) == t) continue;
				while (t.length > 0) {
					t = t.slice(0, t.length - 1);
					t = t.slice(0, t.lastIndexOf("/") + 1);
					if (i.slice(0, t.length) == t) break
				}
			}
			var s = (a[1] || "").match(/boundary="(.*?)"/);
			if (!s) throw new Error("MAD cannot find boundary");
			var l = "--" + (s[1] || "");
			var o = [],
			c = [];
			var f = {
				FileIndex: o,
				FullPaths: c,
			};
			S(f);
			var u, h = 0;
			for (n = 0; n < a.length; ++n) {
				var d = a[n];
				if (d !== l && d !== l + "--") continue;
				if (h++) Ue(f, a.slice(u, n), t);
				u = n
			}
			return f
		}
		function $e(e, r) {
			var t = r || {};
			var a = t.boundary || "SheetJS";
			a = "------=" + a;
			var n = ["MIME-Version: 1.0", 'Content-Type: multipart/related; boundary="' + a.slice(2) + '"', "", "", "", ];
			var i = e.FullPaths[0],
			s = i,
			l = e.FileIndex[0];
			for (var o = 1; o < e.FullPaths.length; ++o) {
				s = e.FullPaths[o].slice(i.length);
				l = e.FileIndex[o];
				if (!l.size || !l.content || s == "Sh33tJ5") continue;
				s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7E-\xFF]/g,
				function(e) {
					return "_x" + e.charCodeAt(0).toString(16) + "_"
				}).replace(/[\u0080-\uFFFF]/g,
				function(e) {
					return "_u" + e.charCodeAt(0).toString(16) + "_"
				});
				var c = l.content;
				var f = _ && Buffer.isBuffer(c) ? c.toString("binary") : H(c);
				var u = 0,
				h = Math.min(1024, f.length),
				d = 0;
				for (var p = 0; p <= h; ++p) if ((d = f.charCodeAt(p)) >= 32 && d < 128)++u;
				var m = u >= (h * 4) / 5;
				n.push(a);
				n.push("Content-Location: " + (t.root || "file:///C:/SheetJS/") + s);
				n.push("Content-Transfer-Encoding: " + (m ? "quoted-printable": "base64"));
				n.push("Content-Type: " + Pe(l, s));
				n.push("");
				n.push(m ? Le(f) : Re(f))
			}
			n.push(a + "--\r\n");
			return n.join("\r\n")
		}
		function We(e) {
			var r = {};
			S(r, e);
			return r
		}
		function je(e, r, t, n) {
			var i = n && n.unsafe;
			if (!i) S(e);
			var s = !i && Qe.find(e, r);
			if (!s) {
				var l = e.FullPaths[0];
				if (r.slice(0, l.length) == l) l = r;
				else {
					if (l.slice(-1) != "/") l += "/";
					l = (l + r).replace("//", "/")
				}
				s = {
					name: a(r),
					type: 2,
				};
				e.FileIndex.push(s);
				e.FullPaths.push(l);
				if (!i) Qe.utils.cfb_gc(e)
			}
			s.content = t;
			s.size = t ? t.length: 0;
			if (n) {
				if (n.CLSID) s.clsid = n.CLSID;
				if (n.mt) s.mt = n.mt;
				if (n.ct) s.ct = n.ct
			}
			return s
		}
		function He(e, r) {
			S(e);
			var t = Qe.find(e, r);
			if (t) for (var a = 0; a < e.FileIndex.length; ++a) if (e.FileIndex[a] == t) {
				e.FileIndex.splice(a, 1);
				e.FullPaths.splice(a, 1);
				return true
			}
			return false
		}
		function Ve(e, r, t) {
			S(e);
			var n = Qe.find(e, r);
			if (n) for (var i = 0; i < e.FileIndex.length; ++i) if (e.FileIndex[i] == n) {
				e.FileIndex[i].name = a(t);
				e.FullPaths[i] = t;
				return true
			}
			return false
		}
		function Xe(e) {
			O(e, true)
		}
		e.find = N;
		e.read = x;
		e.parse = c;
		e.write = V;
		e.writeFile = j;
		e.utils = {
			cfb_new: We,
			cfb_add: je,
			cfb_del: He,
			cfb_mov: Ve,
			cfb_gc: Xe,
			ReadShift: ma,
			CheckField: ka,
			prep_blob: ya,
			bconcat: P,
			use_zlib: G,
			_deflateRaw: ye,
			_inflateRaw: Fe,
			consts: W,
		};
		return e
	})();
	var er;
	function nr(e) {
		if (typeof er !== "undefined") return er.readFileSync(e);
		if (typeof Deno !== "undefined") return Deno.readFileSync(e);
		if (typeof $ !== "undefined" && typeof File !== "undefined" && typeof Folder !== "undefined") try {
			var r = File(e);
			r.open("r");
			r.encoding = "binary";
			var t = r.read();
			r.close();
			return t
		} catch(a) {
			if (!a.message || !a.message.match(/onstruct/)) throw a;
		}
		throw new Error("Cannot access file " + e);
	}
	function ir(e) {
		var r = Object.keys(e),
		t = [];
		for (var a = 0; a < r.length; ++a) if (Object.prototype.hasOwnProperty.call(e, r[a])) t.push(r[a]);
		return t
	}
	function lr(e) {
		var r = [],
		t = ir(e);
		for (var a = 0; a !== t.length; ++a) r[e[t[a]]] = t[a];
		return r
	}
	var fr = Date.UTC(1899, 11, 30, 0, 0, 0);
	var ur = Date.UTC(1899, 11, 31, 0, 0, 0);
	var hr = Date.UTC(1904, 0, 1, 0, 0, 0);
	function dr(e, r) {
		var t = e.getTime();
		var a = (t - fr) / (24 * 60 * 60 * 1e3);
		if (r) {
			a -= 1462;
			return a < -1402 ? a - 1 : a
		}
		return a < 60 ? a - 1 : a
	}
	function pr(e) {
		if (e >= 60 && e < 61) return e;
		var r = new Date();
		r.setTime((e > 60 ? e: e + 1) * 24 * 60 * 60 * 1e3 + fr);
		return r
	}
	function mr(e) {
		var r = 0,
		t = 0,
		a = false;
		var n = e.match(/P([0-9\.]+Y)?([0-9\.]+M)?([0-9\.]+D)?T([0-9\.]+H)?([0-9\.]+M)?([0-9\.]+S)?/);
		if (!n) throw new Error("|" + e + "| is not an ISO8601 Duration");
		for (var i = 1; i != n.length; ++i) {
			if (!n[i]) continue;
			t = 1;
			if (i > 3) a = true;
			switch (n[i].slice(n[i].length - 1)) {
			case "Y":
				throw new Error("Unsupported ISO Duration Field: " + n[i].slice(n[i].length - 1));
			case "D":
				t *= 24;
			case "H":
				t *= 60;
			case "M":
				if (!a) throw new Error("Unsupported ISO Duration Field: M");
				else t *= 60;
			case "S":
				break
			}
			r += t * parseInt(n[i], 10)
		}
		return r
	}
	var vr = /^(\d+):(\d+)(:\d+)?(\.\d+)?$/;
	var gr = /^(\d+)-(\d+)-(\d+)$/;
	var br = /^(\d+)-(\d+)-(\d+)[T ](\d+):(\d+)(:\d+)?(\.\d+)?$/;
	function wr(e, r) {
		if (e instanceof Date) return e;
		var t = e.match(vr);
		if (t) return new Date((r ? hr: ur) + ((parseInt(t[1], 10) * 60 + parseInt(t[2], 10)) * 60 + (t[3] ? parseInt(t[3].slice(1), 10) : 0)) * 1e3 + (t[4] ? parseInt((t[4] + "000").slice(1, 4), 10) : 0));
		t = e.match(gr);
		if (t) return new Date(Date.UTC( + t[1], +t[2] - 1, +t[3], 0, 0, 0, 0));
		t = e.match(br);
		if (t) return new Date(Date.UTC( + t[1], +t[2] - 1, +t[3], +t[4], +t[5], (t[6] && parseInt(t[6].slice(1), 10)) || 0, (t[7] && parseInt(t[7].slice(1), 10)) || 0));
		var a = new Date(e);
		return a
	}
	function kr(e, r) {
		if (_ && Buffer.isBuffer(e)) {
			if (r && T) {
				if (e[0] == 255 && e[1] == 254) return yt(e.slice(2).toString("utf16le"));
				if (e[1] == 254 && e[2] == 255) return yt(d(e.slice(2).toString("binary")))
			}
			return e.toString("binary")
		}
		if (typeof TextDecoder !== "undefined") try {
			if (r) {
				if (e[0] == 255 && e[1] == 254) return yt(new TextDecoder("utf-16le").decode(e.slice(2)));
				if (e[0] == 254 && e[1] == 255) return yt(new TextDecoder("utf-16be").decode(e.slice(2)))
			}
			var t = {
				"€": "",
				"‚": "",
				ƒ: "",
				"„": "",
				"…": "",
				"†": "",
				"‡": "",
				ˆ: "",
				"‰": "",
				Š: "",
				"‹": "",
				Œ: "",
				Ž: "",
				"‘": "",
				"’": "",
				"“": "",
				"”": "",
				"•": "",
				"–": "",
				"—": "",
				"˜": "",
				"™": "",
				š: "",
				"›": "",
				œ: "",
				ž: "",
				Ÿ: "",
			};
			if (Array.isArray(e)) e = new Uint8Array(e);
			return new TextDecoder("latin1").decode(e).replace(/[€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ]/g,
			function(e) {
				return t[e] || e
			})
		} catch(a) {}
		var n = [],
		i = 0;
		try {
			for (i = 0; i < e.length - 65536; i += 65536) n.push(String.fromCharCode.apply(0, e.slice(i, i + 65536)));
			n.push(String.fromCharCode.apply(0, e.slice(i)))
		} catch(a) {
			try {
				for (; i < e.length - 16384; i += 16384) n.push(String.fromCharCode.apply(0, e.slice(i, i + 16384)));
				n.push(String.fromCharCode.apply(0, e.slice(i)))
			} catch(a) {
				for (; i != e.length; ++i) n.push(String.fromCharCode(e[i]))
			}
		}
		return n.join("")
	}
	function yr(e) {
		if (typeof JSON != "undefined" && !Array.isArray(e)) return JSON.parse(JSON.stringify(e));
		if (typeof e != "object" || e == null) return e;
		if (e instanceof Date) return new Date(e.getTime());
		var r = {};
		for (var t in e) if (Object.prototype.hasOwnProperty.call(e, t)) r[t] = yr(e[t]);
		return r
	}
	function xr(e, r) {
		var t = "";
		while (t.length < r) t += e;
		return t
	}
	function Sr(e) {
		var r = Number(e);
		if (!isNaN(r)) return isFinite(r) ? r: NaN;
		if (!/\d/.test(e)) return r;
		var t = 1;
		var a = e.replace(/([\d]),([\d])/g, "$1$2").replace(/[$]/g, "").replace(/[%]/g,
		function() {
			t *= 100;
			return ""
		});
		if (!isNaN((r = Number(a)))) return r / t;
		a = a.replace(/[(](.*)[)]/,
		function(e, r) {
			t = -t;
			return r
		});
		if (!isNaN((r = Number(a)))) return r / t;
		return r
	}
	var Cr = /^(0?\d|1[0-2])(?:|:([0-5]?\d)(?:|(\.\d+)(?:|:([0-5]?\d))|:([0-5]?\d)(|\.\d+)))\s+([ap])m?$/;
	var _r = /^([01]?\d|2[0-3])(?:|:([0-5]?\d)(?:|(\.\d+)(?:|:([0-5]?\d))|:([0-5]?\d)(|\.\d+)))$/;
	var Ar = /^(\d+)-(\d+)-(\d+)[T ](\d+):(\d+)(:\d+)(\.\d+)?[Z]?$/;
	var Tr = new Date("6/9/69 00:00 UTC").valueOf() == -177984e5;
	function Er(e) {
		if (!e[2]) return new Date(Date.UTC(1899, 11, 31, ( + e[1] % 12) + (e[7] == "p" ? 12 : 0), 0, 0, 0));
		if (e[3]) {
			if (e[4]) return new Date(Date.UTC(1899, 11, 31, ( + e[1] % 12) + (e[7] == "p" ? 12 : 0), +e[2], +e[4], parseFloat(e[3]) * 1e3));
			else return new Date(Date.UTC(1899, 11, 31, e[7] == "p" ? 12 : 0, +e[1], +e[2], parseFloat(e[3]) * 1e3))
		} else if (e[5]) return new Date(Date.UTC(1899, 11, 31, ( + e[1] % 12) + (e[7] == "p" ? 12 : 0), +e[2], +e[5], e[6] ? parseFloat(e[6]) * 1e3: 0));
		else return new Date(Date.UTC(1899, 11, 31, ( + e[1] % 12) + (e[7] == "p" ? 12 : 0), +e[2], 0, 0))
	}
	function Fr(e) {
		if (!e[2]) return new Date(Date.UTC(1899, 11, 31, +e[1], 0, 0, 0));
		if (e[3]) {
			if (e[4]) return new Date(Date.UTC(1899, 11, 31, +e[1], +e[2], +e[4], parseFloat(e[3]) * 1e3));
			else return new Date(Date.UTC(1899, 11, 31, 0, +e[1], +e[2], parseFloat(e[3]) * 1e3))
		} else if (e[5]) return new Date(Date.UTC(1899, 11, 31, +e[1], +e[2], +e[5], e[6] ? parseFloat(e[6]) * 1e3: 0));
		else return new Date(Date.UTC(1899, 11, 31, +e[1], +e[2], 0, 0))
	}
	var Dr = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december", ];
	function Or(e) {
		if (Ar.test(e)) return e.indexOf("Z") == -1 ? Ir(new Date(e)) : new Date(e);
		var r = e.toLowerCase();
		var t = r.replace(/\s+/g, " ").trim();
		var a = t.match(Cr);
		if (a) return Er(a);
		a = t.match(_r);
		if (a) return Fr(a);
		a = t.match(br);
		if (a) return new Date(Date.UTC( + a[1], +a[2] - 1, +a[3], +a[4], +a[5], (a[6] && parseInt(a[6].slice(1), 10)) || 0, (a[7] && parseInt(a[7].slice(1), 10)) || 0));
		var n = new Date(Tr && e.indexOf("UTC") == -1 ? e + " UTC": e),
		i = new Date(NaN);
		if (isNaN(o)) return i;
		if (r.match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)) {
			r = r.replace(/[^a-z]/g, "").replace(/([^a-z]|^)[ap]m?([^a-z]|$)/, "");
			if (r.length > 3 && Dr.indexOf(r) == -1) return i
		} else if (r.replace(/[ap]m?/, "").match(/[a-z]/)) return i;
		if (s < 0 || s > 8099 || e.match(/[^-0-9:,\/\\\ ]/)) return i;
		return n
	}
	function Nr(e) {
		return new Date(e.getUTCFullYear(), e.getUTCMonth(), e.getUTCDate(), e.getUTCHours(), e.getUTCMinutes(), e.getUTCSeconds(), e.getUTCMilliseconds())
	}
	function Ir(e) {
		return new Date(Date.UTC(e.getFullYear(), e.getMonth(), e.getDate(), e.getHours(), e.getMinutes(), e.getSeconds(), e.getMilliseconds()))
	}
	function Pr(e) {
		if (!e) return null;
		if (e.content && e.type) return kr(e.content, true);
		if (e.data) return p(e.data);
		if (e.asNodeBuffer && _) return p(e.asNodeBuffer().toString("binary"));
		if (e.asBinary) return p(e.asBinary());
		if (e._data && e._data.getContent) return p(kr(Array.prototype.slice.call(e._data.getContent(), 0)));
		return null
	}
	function Rr(e) {
		if (!e) return null;
		if (e.data) return f(e.data);
		if (e.asNodeBuffer && _) return e.asNodeBuffer();
		if (e._data && e._data.getContent) {
			var r = e._data.getContent();
			if (typeof r == "string") return f(r);
			return Array.prototype.slice.call(r)
		}
		if (e.content && e.type) return e.content;
		return null
	}
	function Lr(e) {
		return e && e.name.slice(-4) === ".bin" ? Rr(e) : Pr(e)
	}
	function Br(e, r) {
		var t = e.FullPaths || ir(e.files);
		var a = r.toLowerCase().replace(/[\/]/g, "\\"),
		n = a.replace(/\\/g, "/");
		for (var i = 0; i < t.length; ++i) {
			var s = t[i].replace(/^Root Entry[\/]/, "").toLowerCase();
			if (a == s || n == s) return e.files ? e.files[t[i]] : e.FileIndex[i]
		}
		return null
	}
	function Ur(e, r) {
		var t = Br(e, r);
		if (t == null) throw new Error("Cannot find file " + r + " in zip");
		return t
	}
	function zr(e, r, t) {
		if (!t) return Lr(Ur(e, r));
		if (!r) return null;
		try {
			return zr(e, r)
		} catch(a) {
			return null
		}
	}
	function $r(e, r, t) {
		if (!t) return Pr(Ur(e, r));
		if (!r) return null;
		try {
			return $r(e, r)
		} catch(a) {
			return null
		}
	}
	function Wr(e, r, t) {
		if (!t) return Rr(Ur(e, r));
		if (!r) return null;
		try {
			return Wr(e, r)
		} catch(a) {
			return null
		}
	}
	function jr(e) {
		var r = e.FullPaths || ir(e.files),
		t = [];
		for (var a = 0; a < r.length; ++a) if (r[a].slice(-1) != "/") t.push(r[a].replace(/^Root Entry[\/]/, ""));
		return t.sort()
	}
	function Hr(e, r, t) {
		if (e.FullPaths) {
			if (typeof t == "string") {
				var a;
				if (_) a = A(t);
				else a = R(t);
				return Qe.utils.cfb_add(e, r, a)
			}
			Qe.utils.cfb_add(e, r, t)
		} else e.file(r, t)
	}
	function Xr(e, r) {
		switch (r.type) {
		case "base64":
			return Qe.read(e, {
				type: "base64",
			});
		case "binary":
			return Qe.read(e, {
				type: "binary",
			});
		case "buffer":
		case "array":
			return Qe.read(e, {
				type: "buffer",
			})
		}
		throw new Error("Unrecognized type " + r.type);
	}
	function Gr(e, r) {
		if (e.charAt(0) == "/") return e.slice(1);
		var t = r.split("/");
		if (r.slice(-1) != "/") t.pop();
		var a = e.split("/");
		while (a.length !== 0) {
			var n = a.shift();
			if (n === "..") t.pop();
			else if (n !== ".") t.push(n)
		}
		return t.join("/")
	}
	var Yr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
	var Jr = /([^"\s?>\/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;
	var Zr = /<[\/\?]?[a-zA-Z0-9:_-]+(?:\s+[^"\s?>\/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^'">\s=]+))*\s*[\/\?]?>/gm,
	Kr = /<[^>]*>/g;
	var qr = Yr.match(Zr) ? Zr: Kr;
	var Qr = /<\w*:/,
	et = /<(\/?)\w+:/;
	function rt(e, r, t) {
		var a = {};
		var n = 0,
		i = 0;
		for (; n !== e.length; ++n) if ((i = e.charCodeAt(n)) === 32 || i === 10 || i === 13) break;
		if (!r) a[0] = e.slice(0, n);
		if (n === e.length) return a;
		var s = e.match(Jr),
		l = 0,
		o = "",
		c = 0,
		f = "",
		u = "",
		h = 1;
		if (s) for (c = 0; c != s.length; ++c) {
			u = s[c];
			for (i = 0; i != u.length; ++i) if (u.charCodeAt(i) === 61) break;
			f = u.slice(0, i).trim();
			while (u.charCodeAt(i + 1) == 32)++i;
			h = (n = u.charCodeAt(i + 1)) == 34 || n == 39 ? 1 : 0;
			o = u.slice(i + 1 + h, u.length - h);
			for (l = 0; l != f.length; ++l) if (f.charCodeAt(l) === 58) break;
			if (l === f.length) {
				if (f.indexOf("_") > 0) f = f.slice(0, f.indexOf("_"));
				a[f] = o;
				if (!t) a[f.toLowerCase()] = o
			} else {
				var d = (l === 5 && f.slice(0, 5) === "xmlns" ? "xmlns": "") + f.slice(l + 1);
				if (a[d] && f.slice(l - 3, l) == "ext") continue;
				a[d] = o;
				if (!t) a[d.toLowerCase()] = o
			}
		}
		return a
	}
	function tt(e) {
		return e.replace(et, "<$1")
	}
	var at = {
		"&quot;": '"',
		"&apos;": "'",
		"&gt;": ">",
		"&lt;": "<",
		"&amp;": "&",
	};
	var nt = lr(at);
	var it = (function() {
		var e = /&(?:quot|apos|gt|lt|amp|#x?([\da-fA-F]+));/gi,
		r = /_x([\da-fA-F]{4})_/gi;
		function t(a) {
			var n = a + "",
			i = n.indexOf("<![CDATA[");
			if (i == -1) return n.replace(e,
			function(e, r) {
				return (at[e] || String.fromCharCode(parseInt(r, e.indexOf("x") > -1 ? 16 : 10)) || e)
			}).replace(r,
			function(e, r) {
				return String.fromCharCode(parseInt(r, 16))
			});
			var s = n.indexOf("]]>");
			return t(n.slice(0, i)) + n.slice(i + 9, s) + t(n.slice(s + 3))
		}
		return function a(e, r) {
			var a = t(e);
			return r ? a.replace(/\r\n/g, "\n") : a
		}
	})();
	var st = /[&<>'"]/g,
	lt = /[\u0000-\u0008\u000b-\u001f\uFFFE-\uFFFF]/g;
	function ot(e) {
		var r = e + "";
		return r.replace(st,
		function(e) {
			return nt[e]
		}).replace(lt,
		function(e) {
			return "_x" + ("000" + e.charCodeAt(0).toString(16)).slice(-4) + "_"
		})
	}
	var ft = /[\u0000-\u001f]/g;
	function ut(e) {
		var r = e + "";
		return r.replace(st,
		function(e) {
			return nt[e]
		}).replace(/\n/g, "<br/>").replace(ft,
		function(e) {
			return "&#x" + ("000" + e.charCodeAt(0).toString(16)).slice(-4) + ";"
		})
	}
	function mt(e) {
		switch (e) {
		case 1:
		case true:
		case "1":
		case "true":
			return true;
		case 0:
		case false:
		case "0":
		case "false":
			return false
		}
		return false
	}
	function vt(e) {
		var r = "",
		t = 0,
		a = 0,
		n = 0,
		i = 0,
		s = 0,
		l = 0;
		while (t < e.length) {
			a = e.charCodeAt(t++);
			if (a < 128) {
				r += String.fromCharCode(a);
				continue
			}
			n = e.charCodeAt(t++);
			if (a > 191 && a < 224) {
				s = (a & 31) << 6;
				s |= n & 63;
				r += String.fromCharCode(s);
				continue
			}
			i = e.charCodeAt(t++);
			if (a < 240) {
				r += String.fromCharCode(((a & 15) << 12) | ((n & 63) << 6) | (i & 63));
				continue
			}
			s = e.charCodeAt(t++);
			l = (((a & 7) << 18) | ((n & 63) << 12) | ((i & 63) << 6) | (s & 63)) - 65536;
			r += String.fromCharCode(55296 + ((l >>> 10) & 1023));
			r += String.fromCharCode(56320 + (l & 1023))
		}
		return r
	}
	function gt(e) {
		var r = E(2 * e.length),
		t,
		a,
		n = 1,
		i = 0,
		s = 0,
		l;
		for (a = 0; a < e.length; a += n) {
			n = 1;
			if ((l = e.charCodeAt(a)) < 128) t = l;
			else if (l < 224) {
				t = (l & 31) * 64 + (e.charCodeAt(a + 1) & 63);
				n = 2
			} else if (l < 240) {
				t = (l & 15) * 4096 + (e.charCodeAt(a + 1) & 63) * 64 + (e.charCodeAt(a + 2) & 63);
				n = 3
			} else {
				n = 4;
				t = (l & 7) * 262144 + (e.charCodeAt(a + 1) & 63) * 4096 + (e.charCodeAt(a + 2) & 63) * 64 + (e.charCodeAt(a + 3) & 63);
				t -= 65536;
				s = 55296 + ((t >>> 10) & 1023);
				t = 56320 + (t & 1023)
			}
			if (s !== 0) {
				r[i++] = s & 255;
				r[i++] = s >>> 8;
				s = 0
			}
			r[i++] = t % 256;
			r[i++] = t >>> 8
		}
		return r.slice(0, i).toString("ucs2")
	}
	function bt(e) {
		return A(e, "binary").toString("utf8")
	}
	var wt = "foo bar bazâð£";
	var kt = (_ && ((bt(wt) == vt(wt) && bt) || (gt(wt) == vt(wt) && gt))) || vt;
	var yt = _ ?
	function(e) {
		return A(e, "utf8").toString("binary")
	}: function(e) {
		var r = [],
		t = 0,
		a = 0,
		n = 0;
		while (t < e.length) {
			a = e.charCodeAt(t++);
			switch (true) {
			case a < 128 : r.push(String.fromCharCode(a));
				break;
			case a < 2048 : r.push(String.fromCharCode(192 + (a >> 6)));
				r.push(String.fromCharCode(128 + (a & 63)));
				break;
			case a >= 55296 && a < 57344 : a -= 55296;
				n = e.charCodeAt(t++) - 56320 + (a << 10);
				r.push(String.fromCharCode(240 + ((n >> 18) & 7)));
				r.push(String.fromCharCode(144 + ((n >> 12) & 63)));
				r.push(String.fromCharCode(128 + ((n >> 6) & 63)));
				r.push(String.fromCharCode(128 + (n & 63)));
				break;
			default:
				r.push(String.fromCharCode(224 + (a >> 12)));
				r.push(String.fromCharCode(128 + ((a >> 6) & 63)));
				r.push(String.fromCharCode(128 + (a & 63)))
			}
		}
		return r.join("")
	};
	var xt = (function() {
		var e = {};
		return function r(t, a) {
			var n = t + "|" + (a || "");
			if (e[n]) return e[n];
			return (e[n] = new RegExp("<(?:\\w+:)?" + t + '(?: xml:space="preserve")?(?:[^>]*)>([\\s\\S]*?)</(?:\\w+:)?' + t + ">", a || ""))
		}
	})();
	var St = (function() {
		var e = [["nbsp", " "], ["middot", "·"], ["quot", '"'], ["apos", "'"], ["gt", ">"], ["lt", "<"], ["amp", "&"], ].map(function(e) {
			return [new RegExp("&" + e[0] + ";", "ig"), e[1]]
		});
		return function r(t) {
			var a = t.replace(/^[\t\n\r ]+/, "").replace(/[\t\n\r ]+$/, "").replace(/>\s+/g, ">").replace(/\s+</g, "<").replace(/[\t\n\r ]+/g, " ").replace(/<\s*[bB][rR]\s*\/?>/g, "\n").replace(/<[^>]*>/g, "");
			for (var n = 0; n < e.length; ++n) a = a.replace(e[n][0], e[n][1]);
			return a
		}
	})();
	var Ct = (function() {
		var e = {};
		return function r(t) {
			if (e[t] !== undefined) return e[t];
			return (e[t] = new RegExp("<(?:vt:)?" + t + ">([\\s\\S]*?)</(?:vt:)?" + t + ">", "g"))
		}
	})();
	var _t = /<\/?(?:vt:)?variant>/g,
	At = /<(?:vt:)([^>]*)>([\s\S]*)</;
	function Tt(e, r) {
		var t = rt(e);
		var a = e.match(Ct(t.baseType)) || [];
		var n = [];
		if (a.length != t.size) {
			if (r.WTF) throw new Error("unexpected vector length " + a.length + " != " + t.size);
			return n
		}
		a.forEach(function(e) {
			var r = e.replace(_t, "").match(At);
			if (r) n.push({
				v: kt(r[2]),
				t: r[1],
			})
		});
		return n
	}
	var Et = /(^\s|\s$|\n)/;
	function Dt(e) {
		return ir(e).map(function(r) {
			return " " + r + '="' + e[r] + '"'
		}).join("")
	}
	function Ot(e, r, t) {
		return ("<" + e + (t != null ? Dt(t) : "") + (r != null ? (r.match(Et) ? ' xml:space="preserve"': "") + ">" + r + "</" + e: "/") + ">")
	}
	function It(e) {
		if (_ && Buffer.isBuffer(e)) return e.toString("utf8");
		if (typeof e === "string") return e;
		if (typeof Uint8Array !== "undefined" && e instanceof Uint8Array) return kt(M(I(e)));
		throw new Error("Bad input format: expected Buffer or string");
	}
	var Pt = /<(\/?)([^\s?><!\/:]*:|)([^\s?<>:\/]+)(?:[\s?:\/](?:[^>=]|="[^"]*?")*)?>/gm;
	var Rt = {
		CORE_PROPS: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
		CUST_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
		EXT_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
		CT: "http://schemas.openxmlformats.org/package/2006/content-types",
		RELS: "http://schemas.openxmlformats.org/package/2006/relationships",
		TCMNT: "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
		dc: "http://purl.org/dc/elements/1.1/",
		dcterms: "http://purl.org/dc/terms/",
		dcmitype: "http://purl.org/dc/dcmitype/",
		mx: "http://schemas.microsoft.com/office/mac/excel/2008/main",
		r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		sjs: "http://schemas.openxmlformats.org/package/2006/sheetjs/core-properties",
		vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
		xsi: "http://www.w3.org/2001/XMLSchema-instance",
		xsd: "http://www.w3.org/2001/XMLSchema",
	};
	var Lt = ["http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main", "http://schemas.microsoft.com/office/excel/2006/main", "http://schemas.microsoft.com/office/excel/2006/2", ];
	function Ut(e, r) {
		var t = 1 - 2 * (e[r + 7] >>> 7);
		var a = ((e[r + 7] & 127) << 4) + ((e[r + 6] >>> 4) & 15);
		var n = e[r + 6] & 15;
		for (var i = 5; i >= 0; --i) n = n * 256 + e[r + i];
		if (a == 2047) return n == 0 ? t * Infinity: NaN;
		if (a == 0) a = -1022;
		else {
			a -= 1023;
			n += Math.pow(2, 52)
		}
		return t * Math.pow(2, a - 52) * n
	}
	function zt(e, r, t) {
		var a = (r < 0 || 1 / r == -Infinity ? 1 : 0) << 7,
		n = 0,
		i = 0;
		var s = a ? -r: r;
		if (!isFinite(s)) {
			n = 2047;
			i = isNaN(r) ? 26985 : 0
		} else if (s == 0) n = i = 0;
		else {
			n = Math.floor(Math.log(s) / Math.LN2);
			i = s * Math.pow(2, 52 - n);
			if (n <= -1023 && (!isFinite(i) || i < Math.pow(2, 52))) {
				n = -1022
			} else {
				i -= Math.pow(2, 52);
				n += 1023
			}
		}
		for (var l = 0; l <= 5; ++l, i /= 256) e[t + l] = i & 255;
		e[t + 6] = ((n & 15) << 4) | (i & 15);
		e[t + 7] = (n >> 4) | a
	}
	var $t = function(e) {
		var r = [],
		t = 10240;
		for (var a = 0; a < e[0].length; ++a) if (e[0][a]) for (var n = 0,
		i = e[0][a].length; n < i; n += t) r.push.apply(r, e[0][a].slice(n, n + t));
		return r
	};
	var Wt = _ ?
	function(e) {
		return e[0].length > 0 && Buffer.isBuffer(e[0][0]) ? Buffer.concat(e[0].map(function(e) {
			return Buffer.isBuffer(e) ? e: A(e)
		})) : $t(e)
	}: $t;
	var jt = function(e, r, t) {
		var a = [];
		for (var n = r; n < t; n += 2) a.push(String.fromCharCode(fa(e, n)));
		return a.join("").replace(L, "")
	};
	var Ht = _ ?
	function(e, r, t) {
		if (!Buffer.isBuffer(e) || !T) return jt(e, r, t);
		return e.toString("utf16le", r, t).replace(L, "")
	}: jt;
	var Vt = function(e, r, t) {
		var a = [];
		for (var n = r; n < r + t; ++n) a.push(("0" + e[n].toString(16)).slice(-2));
		return a.join("")
	};
	var Xt = _ ?
	function(e, r, t) {
		return Buffer.isBuffer(e) ? e.toString("hex", r, r + t) : Vt(e, r, t)
	}: Vt;
	var Gt = function(e, r, t) {
		var a = [];
		for (var n = r; n < t; n++) a.push(String.fromCharCode(ca(e, n)));
		return a.join("")
	};
	var Yt = _ ?
	function Bc(e, r, t) {
		return Buffer.isBuffer(e) ? e.toString("utf8", r, t) : Gt(e, r, t)
	}: Gt;
	var Jt = function(e, r) {
		var t = ha(e, r);
		return t > 0 ? Yt(e, r + 4, r + 4 + t - 1) : ""
	};
	var Zt = Jt;
	var Kt = function(e, r) {
		var t = ha(e, r);
		return t > 0 ? Yt(e, r + 4, r + 4 + t - 1) : ""
	};
	var qt = Kt;
	var Qt = function(e, r) {
		var t = 2 * ha(e, r);
		return t > 0 ? Yt(e, r + 4, r + 4 + t - 1) : ""
	};
	var ea = Qt;
	var ra = function Uc(e, r) {
		var t = ha(e, r);
		return t > 0 ? Ht(e, r + 4, r + 4 + t) : ""
	};
	var ta = ra;
	var aa = function(e, r) {
		var t = ha(e, r);
		return t > 0 ? Yt(e, r + 4, r + 4 + t) : ""
	};
	var na = aa;
	var ia = function(e, r) {
		return Ut(e, r)
	};
	var sa = ia;
	var la = function zc(e) {
		return (Array.isArray(e) || (typeof Uint8Array !== "undefined" && e instanceof Uint8Array))
	};
	if (_) {
		Zt = function $c(e, r) {
			if (!Buffer.isBuffer(e)) return Jt(e, r);
			var t = e.readUInt32LE(r);
			return t > 0 ? e.toString("utf8", r + 4, r + 4 + t - 1) : ""
		};
		qt = function Wc(e, r) {
			if (!Buffer.isBuffer(e)) return Kt(e, r);
			var t = e.readUInt32LE(r);
			return t > 0 ? e.toString("utf8", r + 4, r + 4 + t - 1) : ""
		};
		ea = function jc(e, r) {
			if (!Buffer.isBuffer(e) || !T) return Qt(e, r);
			var t = 2 * e.readUInt32LE(r);
			return e.toString("utf16le", r + 4, r + 4 + t - 1)
		};
		ta = function Hc(e, r) {
			if (!Buffer.isBuffer(e) || !T) return ra(e, r);
			var t = e.readUInt32LE(r);
			return e.toString("utf16le", r + 4, r + 4 + t)
		};
		na = function Vc(e, r) {
			if (!Buffer.isBuffer(e)) return aa(e, r);
			var t = e.readUInt32LE(r);
			return e.toString("utf8", r + 4, r + 4 + t)
		};
		sa = function Xc(e, r) {
			if (Buffer.isBuffer(e)) return e.readDoubleLE(r);
			return ia(e, r)
		};
		la = function Gc(e) {
			return (Buffer.isBuffer(e) || Array.isArray(e) || (typeof Uint8Array !== "undefined" && e instanceof Uint8Array))
		}
	}
	function oa() {
		Ht = function(e, r, t) {
			return a.utils.decode(1200, e.slice(r, t)).replace(L, "")
		};
		Yt = function(e, r, t) {
			return a.utils.decode(65001, e.slice(r, t))
		};
		Zt = function(e, r) {
			var n = ha(e, r);
			return n > 0 ? a.utils.decode(t, e.slice(r + 4, r + 4 + n - 1)) : ""
		};
		qt = function(e, t) {
			var n = ha(e, t);
			return n > 0 ? a.utils.decode(r, e.slice(t + 4, t + 4 + n - 1)) : ""
		};
		ea = function(e, r) {
			var t = 2 * ha(e, r);
			return t > 0 ? a.utils.decode(1200, e.slice(r + 4, r + 4 + t - 1)) : ""
		};
		ta = function(e, r) {
			var t = ha(e, r);
			return t > 0 ? a.utils.decode(1200, e.slice(r + 4, r + 4 + t)) : ""
		};
		na = function(e, r) {
			var t = ha(e, r);
			return t > 0 ? a.utils.decode(65001, e.slice(r + 4, r + 4 + t)) : ""
		}
	}
	if (typeof a !== "undefined") oa();
	var ca = function(e, r) {
		return e[r]
	};
	var fa = function(e, r) {
		return e[r + 1] * (1 << 8) + e[r]
	};
	var ua = function(e, r) {
		var t = e[r + 1] * (1 << 8) + e[r];
		return t < 32768 ? t: (65535 - t + 1) * -1
	};
	var ha = function(e, r) {
		return e[r + 3] * (1 << 24) + (e[r + 2] << 16) + (e[r + 1] << 8) + e[r]
	};
	var da = function(e, r) {
		return (e[r + 3] << 24) | (e[r + 2] << 16) | (e[r + 1] << 8) | e[r]
	};
	var pa = function(e, r) {
		return (e[r] << 24) | (e[r + 1] << 16) | (e[r + 2] << 8) | e[r + 3]
	};
	function ma(e, t) {
		var n = "",
		i, s, l = [],
		o,
		c,
		f,
		u;
		switch (t) {
		case "dbcs":
			u = this.l;
			if (_ && Buffer.isBuffer(this) && T) n = this.slice(this.l, this.l + 2 * e).toString("utf16le");
			else for (f = 0; f < e; ++f) {
				n += String.fromCharCode(fa(this, u));
				u += 2
			}
			e *= 2;
			break;
		case "utf8":
			n = Yt(this, this.l, this.l + e);
			break;
		case "utf16le":
			e *= 2;
			n = Ht(this, this.l, this.l + e);
			break;
		case "wstr":
			if (typeof a !== "undefined") n = a.utils.decode(r, this.slice(this.l, this.l + 2 * e));
			else return ma.call(this, e, "dbcs");
			e = 2 * e;
			break;
		case "lpstr-ansi":
			n = Zt(this, this.l);
			e = 4 + ha(this, this.l);
			break;
		case "lpstr-cp":
			n = qt(this, this.l);
			e = 4 + ha(this, this.l);
			break;
		case "lpwstr":
			n = ea(this, this.l);
			e = 4 + 2 * ha(this, this.l);
			break;
		case "lpp4":
			e = 4 + ha(this, this.l);
			n = ta(this, this.l);
			if (e & 2) e += 2;
			break;
		case "8lpp4":
			e = 4 + ha(this, this.l);
			n = na(this, this.l);
			if (e & 3) e += 4 - (e & 3);
			break;
		case "cstr":
			e = 0;
			n = "";
			while ((o = ca(this, this.l + e++)) !== 0) l.push(m(o));
			n = l.join("");
			break;
		case "_wstr":
			e = 0;
			n = "";
			while ((o = fa(this, this.l + e)) !== 0) {
				l.push(m(o));
				e += 2
			}
			e += 2;
			n = l.join("");
			break;
		case "dbcs-cont":
			n = "";
			u = this.l;
			for (f = 0; f < e; ++f) {
				if (this.lens && this.lens.indexOf(u) !== -1) {
					o = ca(this, u);
					this.l = u + 1;
					c = ma.call(this, e - f, o ? "dbcs-cont": "sbcs-cont");
					return l.join("") + c
				}
				l.push(m(fa(this, u)));
				u += 2
			}
			n = l.join("");
			e *= 2;
			break;
		case "cpstr":
			if (typeof a !== "undefined") {
				n = a.utils.decode(r, this.slice(this.l, this.l + e));
				break
			}
		case "sbcs-cont":
			n = "";
			u = this.l;
			for (f = 0; f != e; ++f) {
				if (this.lens && this.lens.indexOf(u) !== -1) {
					o = ca(this, u);
					this.l = u + 1;
					c = ma.call(this, e - f, o ? "dbcs-cont": "sbcs-cont");
					return l.join("") + c
				}
				l.push(m(ca(this, u)));
				u += 1
			}
			n = l.join("");
			break;
		default:
			switch (e) {
			case 1:
				i = ca(this, this.l);
				this.l++;
				return i;
			case 2:
				i = (t === "i" ? ua: fa)(this, this.l);
				this.l += 2;
				return i;
			case 4:
			case - 4 : if (t === "i" || (this[this.l + 3] & 128) === 0) {
					i = (e > 0 ? da: pa)(this, this.l);
					this.l += 4;
					return i
				} else {
					s = ha(this, this.l);
					this.l += 4
				}
				return s;
			case 8:
			case - 8 : if (t === "f") {
					if (e == 8) s = sa(this, this.l);
					else s = sa([this[this.l + 7], this[this.l + 6], this[this.l + 5], this[this.l + 4], this[this.l + 3], this[this.l + 2], this[this.l + 1], this[this.l + 0], ], 0);
					this.l += 8;
					return s
				} else e = 8;
			case 16:
				n = Xt(this, this.l, e);
				break
			}
		}
		this.l += e;
		return n
	}
	var va = function(e, r, t) {
		e[t] = r & 255;
		e[t + 1] = (r >>> 8) & 255;
		e[t + 2] = (r >>> 16) & 255;
		e[t + 3] = (r >>> 24) & 255
	};
	var ga = function(e, r, t) {
		e[t] = r & 255;
		e[t + 1] = (r >> 8) & 255;
		e[t + 2] = (r >> 16) & 255;
		e[t + 3] = (r >> 24) & 255
	};
	var ba = function(e, r, t) {
		e[t] = r & 255;
		e[t + 1] = (r >>> 8) & 255
	};
	function wa(e, n, i) {
		var s = 0,
		l = 0;
		if (i === "dbcs") {
			for (l = 0; l != n.length; ++l) ba(this, n.charCodeAt(l), this.l + 2 * l);
			s = 2 * n.length
		} else if (i === "sbcs" || i == "cpstr") {
			if (typeof a !== "undefined" && t == 874) {
				for (l = 0; l != n.length; ++l) {
					var o = a.utils.encode(t, n.charAt(l));
					this[this.l + l] = o[0]
				}
				s = n.length
			} else if (typeof a !== "undefined" && i == "cpstr") {
				o = a.utils.encode(r, n);
				if (o.length == n.length) for (l = 0; l < n.length; ++l) if (o[l] == 0 && n.charCodeAt(l) != 0) o[l] = 95;
				if (o.length == 2 * n.length) for (l = 0; l < n.length; ++l) if (o[2 * l] == 0 && o[2 * l + 1] == 0 && n.charCodeAt(l) != 0) o[2 * l] = 95;
				for (l = 0; l < o.length; ++l) this[this.l + l] = o[l];
				s = o.length
			} else {
				n = n.replace(/[^\x00-\x7F]/g, "_");
				for (l = 0; l != n.length; ++l) this[this.l + l] = n.charCodeAt(l) & 255;
				s = n.length
			}
		} else if (i === "hex") {
			for (; l < e; ++l) {
				this[this.l++] = parseInt(n.slice(2 * l, 2 * l + 2), 16) || 0
			}
			return this
		} else if (i === "utf16le") {
			var c = Math.min(this.l + e, this.length);
			for (l = 0; l < Math.min(n.length, e); ++l) {
				var f = n.charCodeAt(l);
				this[this.l++] = f & 255;
				this[this.l++] = f >> 8
			}
			while (this.l < c) this[this.l++] = 0;
			return this
		} else switch (e) {
		case 1:
			s = 1;
			this[this.l] = n & 255;
			break;
		case 2:
			s = 2;
			this[this.l] = n & 255;
			n >>>= 8;
			this[this.l + 1] = n & 255;
			break;
		case 3:
			s = 3;
			this[this.l] = n & 255;
			n >>>= 8;
			this[this.l + 1] = n & 255;
			n >>>= 8;
			this[this.l + 2] = n & 255;
			break;
		case 4:
			s = 4;
			va(this, n, this.l);
			break;
		case 8:
			s = 8;
			if (i === "f") {
				zt(this, n, this.l);
				break
			}
		case 16:
			break;
		case - 4 : s = 4;
			ga(this, n, this.l);
			break
		}
		this.l += s;
		return this
	}
	function ka(e, r) {
		var t = Xt(this, this.l, e.length >> 1);
		if (t !== e) throw new Error(r + "Expected " + e + " saw " + t);
		this.l += e.length >> 1
	}
	function ya(e, r) {
		e.l = r;
		e._R = ma;
		e.chk = ka;
		e._W = wa
	}
	function Sa(e) {
		var r = E(e);
		ya(r, 0);
		return r
	}
	function Ca(e, r, t) {
		if (!e) return;
		var a, n, i;
		ya(e, e.l || 0);
		var s = e.length,
		l = 0,
		o = 0;
		while (e.l < s) {
			l = e._R(1);
			if (l & 128) l = (l & 127) + ((e._R(1) & 127) << 7);
			var c = XLSBRecordEnum[l] || XLSBRecordEnum[65535];
			a = e._R(1);
			i = a & 127;
			for (n = 1; n < 4 && a & 128; ++n) i += ((a = e._R(1)) & 127) << (7 * n);
			o = e.l + i;
			var f = c.f && c.f(e, i, t);
			e.l = o;
			if (r(f, c, l)) return
		}
	}
	function _a() {
		var e = [],
		r = _ ? 256 : 2048;
		var t = function o(e) {
			var r = Sa(e);
			ya(r, 0);
			return r
		};
		var a = t(r);
		var n = function c() {
			if (!a) return;
			if (a.l) {
				if (a.length > a.l) {
					a = a.slice(0, a.l);
					a.l = a.length
				}
				if (a.length > 0) e.push(a)
			}
			a = null
		};
		var i = function f(e) {
			if (a && e < a.length - a.l) return a;
			n();
			return (a = t(Math.max(e + 1, r)))
		};
		var s = function u() {
			n();
			return P(e)
		};
		var l = function h(e) {
			n();
			a = e;
			if (a.l == null) a.l = a.length;
			i(r)
		};
		return {
			next: i,
			push: l,
			end: s,
			_bufs: e,
		}
	}
	function Oa(e) {
		return parseInt(Ia(e), 10) - 1
	}
	function Ma(e) {
		return "" + (e + 1)
	}
	function Ia(e) {
		return e.replace(/\$(\d+)$/, "$1")
	}
	function Pa(e) {
		var r = Ba(e),
		t = 0,
		a = 0;
		for (; a !== r.length; ++a) t = 26 * t + r.charCodeAt(a) - 64;
		return t - 1
	}
	function Ra(e) {
		if (e < 0) throw new Error("invalid column " + e);
		var r = "";
		for (++e; e; e = Math.floor((e - 1) / 26)) r = String.fromCharCode(((e - 1) % 26) + 65) + r;
		return r
	}
	function Ba(e) {
		return e.replace(/^\$([A-Z])/, "$1")
	}
	function Ua(e) {
		return e.replace(/(\$?[A-Z]*)(\$?\d*)/, "$1,$2").split(",")
	}
	function za(e) {
		var r = 0,
		t = 0;
		for (var a = 0; a < e.length; ++a) {
			var n = e.charCodeAt(a);
			if (n >= 48 && n <= 57) r = 10 * r + (n - 48);
			else if (n >= 65 && n <= 90) t = 26 * t + (n - 64)
		}
		return {
			c: t - 1,
			r: r - 1,
		}
	}
	function $a(e) {
		var r = e.c + 1;
		var t = "";
		for (; r; r = ((r - 1) / 26) | 0) t = String.fromCharCode(((r - 1) % 26) + 65) + t;
		return t + (e.r + 1)
	}
	function Wa(e) {
		var r = e.indexOf(":");
		if (r == -1) return {
			s: za(e),
			e: za(e),
		};
		return {
			s: za(e.slice(0, r)),
			e: za(e.slice(r + 1)),
		}
	}
	function ja(e, r) {
		if (typeof r === "undefined" || typeof r === "number") {
			return ja(e.s, e.e)
		}
		if (typeof e !== "string") e = $a(e);
		if (typeof r !== "string") r = $a(r);
		return e == r ? e: e + ":" + r
	}
	function Xa(e) {
		var r = {
			s: {
				c: 0,
				r: 0,
			},
			e: {
				c: 0,
				r: 0,
			},
		};
		var t = 0,
		a = 0,
		n = 0;
		var i = e.length;
		for (t = 0; a < i; ++a) {
			if ((n = e.charCodeAt(a) - 64) < 1 || n > 26) break;
			t = 26 * t + n
		}
		r.s.c = --t;
		for (t = 0; a < i; ++a) {
			if ((n = e.charCodeAt(a) - 48) < 0 || n > 9) break;
			t = 10 * t + n
		}
		r.s.r = --t;
		if (a === i || n != 10) {
			r.e.c = r.s.c;
			r.e.r = r.s.r;
			return r
		}++a;
		for (t = 0; a != i; ++a) {
			if ((n = e.charCodeAt(a) - 64) < 1 || n > 26) break;
			t = 26 * t + n
		}
		r.e.c = --t;
		for (t = 0; a != i; ++a) {
			if ((n = e.charCodeAt(a) - 48) < 0 || n > 9) break;
			t = 10 * t + n
		}
		r.e.r = --t;
		return r
	}
	function Ga(e, r) {
		var t = e.t == "d" && r instanceof Date;
		if (e.z != null) try {
			return (e.w = $e(e.z, t ? dr(r) : r))
		} catch(a) {}
		try {
			return (e.w = $e((e.XF || {}).numFmtId || (t ? 14 : 0), t ? dr(r) : r))
		} catch(a) {
			return "" + r
		}
	}
	function Ya(e, r, t) {
		if (e == null || e.t == null || e.t == "z") return "";
		if (e.w !== undefined) return e.w;
		if (e.t == "d" && !e.z && t && t.dateNF) e.z = t.dateNF;
		if (e.t == "e") return kn[e.v] || e.v;
		if (r == undefined) return Ga(e, e.v);
		return Ga(e, r)
	}
	function Ja(e, r) {
		var t = r && r.sheet ? r.sheet: "Sheet1";
		var a = {};
		a[t] = e;
		return {
			SheetNames: [t],
			Sheets: a,
		}
	}
	function Za(e) {
		var r = {};
		var t = e || {};
		if (t.dense) r["!data"] = [];
		return r
	}
	function Ka(e, r, t) {
		var a = t || {};
		var n = e ? e["!data"] != null: a.dense;
		if (b != null && n == null) n = b;
		var i = e || {};
		if (n && !i["!data"]) i["!data"] = [];
		var s = 0,
		l = 0;
		if (i && a.origin != null) {
			if (typeof a.origin == "number") s = a.origin;
			else {
				var o = typeof a.origin == "string" ? za(a.origin) : a.origin;
				s = o.r;
				l = o.c
			}
			if (!i["!ref"]) i["!ref"] = "A1:A1"
		}
		var c = {
			s: {
				c: 1e7,
				r: 1e7,
			},
			e: {
				c: 0,
				r: 0,
			},
		};
		if (i["!ref"]) {
			var f = Xa(i["!ref"]);
			c.s.c = f.s.c;
			c.s.r = f.s.r;
			c.e.c = Math.max(c.e.c, f.e.c);
			c.e.r = Math.max(c.e.r, f.e.r);
			if (s == -1) c.e.r = s = f.e.r + 1
		}
		var u = [];
		for (var h = 0; h != r.length; ++h) {
			if (!r[h]) continue;
			if (!Array.isArray(r[h])) throw new Error("aoa_to_sheet expects an array of arrays");
			var d = s + h,
			p = "" + (d + 1);
			if (n) {
				if (!i["!data"][d]) i["!data"][d] = [];
				u = i["!data"][d]
			}
			for (var m = 0; m != r[h].length; ++m) {
				if (typeof r[h][m] === "undefined") continue;
				var v = {
					v: r[h][m],
				};
				var g = l + m;
				if (c.s.r > d) c.s.r = d;
				if (c.s.c > g) c.s.c = g;
				if (c.e.r < d) c.e.r = d;
				if (c.e.c < g) c.e.c = g;
				if (r[h][m] && typeof r[h][m] === "object" && !Array.isArray(r[h][m]) && !(r[h][m] instanceof Date)) v = r[h][m];
				else {
					if (Array.isArray(v.v)) {
						v.f = r[h][m][1];
						v.v = v.v[0]
					}
					if (v.v === null) {
						if (v.f) v.t = "n";
						else if (a.nullError) {
							v.t = "e";
							v.v = 0
						} else if (!a.sheetStubs) continue;
						else v.t = "z"
					} else if (typeof v.v === "number") v.t = "n";
					else if (typeof v.v === "boolean") v.t = "b";
					else if (v.v instanceof Date) {
						v.z = a.dateNF || q[14];
						if (!a.UTC) v.v = Ir(v.v);
						if (a.cellDates) {
							v.t = "d";
							v.w = $e(v.z, dr(v.v, a.date1904))
						} else {
							v.t = "n";
							v.v = dr(v.v, a.date1904);
							v.w = $e(v.z, v.v)
						}
					} else v.t = "s"
				}
				if (n) {
					if (u[g] && u[g].z) v.z = u[g].z;
					u[g] = v
				} else {
					var w = Ra(g) + p;
					if (i[w] && i[w].z) v.z = i[w].z;
					i[w] = v
				}
			}
		}
		if (c.s.c < 1e7) i["!ref"] = ja(c);
		return i
	}
	function qa(e, r) {
		return Ka(null, e, r)
	}
	function gn(e) {
		return e.map(function(e) {
			return [(e >> 16) & 255, (e >> 8) & 255, e & 255]
		})
	}
	var bn = gn([0, 16777215, 16711680, 65280, 255, 16776960, 16711935, 65535, 0, 16777215, 16711680, 65280, 255, 16776960, 16711935, 65535, 8388608, 32768, 128, 8421376, 8388736, 32896, 12632256, 8421504, 10066431, 10040166, 16777164, 13434879, 6684774, 16744576, 26316, 13421823, 128, 16711935, 16776960, 65535, 8388736, 8388608, 32896, 255, 52479, 13434879, 13434828, 16777113, 10079487, 16751052, 13408767, 16764057, 3368703, 3394764, 10079232, 16763904, 16750848, 16737792, 6710937, 9868950, 13158, 3381606, 13056, 3355392, 10040064, 10040166, 3355545, 3355443, 0, 16777215, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ]);
	var wn = yr(bn);
	var kn = {
		0 : "#NULL!",
		7 : "#DIV/0!",
		15 : "#VALUE!",
		23 : "#REF!",
		29 : "#NAME?",
		36 : "#NUM!",
		42 : "#N/A",
		43 : "#GETTING_DATA",
		255 : "#WTF?",
	};
	var yn = {
		"#NULL!": 0,
		"#DIV/0!": 7,
		"#VALUE!": 15,
		"#REF!": 23,
		"#NAME?": 29,
		"#NUM!": 36,
		"#N/A": 42,
		"#GETTING_DATA": 43,
		"#WTF?": 255,
	};
	var Sn = {
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks",
		"application/vnd.ms-excel.sheet.macroEnabled.main+xml": "workbooks",
		"application/vnd.ms-excel.sheet.binary.macroEnabled.main": "workbooks",
		"application/vnd.ms-excel.addin.macroEnabled.main+xml": "workbooks",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "workbooks",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml": "sheets",
		"application/vnd.ms-excel.worksheet": "sheets",
		"application/vnd.ms-excel.binIndexWs": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml": "charts",
		"application/vnd.ms-excel.chartsheet": "charts",
		"application/vnd.ms-excel.macrosheet+xml": "macros",
		"application/vnd.ms-excel.macrosheet": "macros",
		"application/vnd.ms-excel.intlmacrosheet": "TODO",
		"application/vnd.ms-excel.binIndexMs": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml": "dialogs",
		"application/vnd.ms-excel.dialogsheet": "dialogs",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "strs",
		"application/vnd.ms-excel.sharedStrings": "strs",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml": "styles",
		"application/vnd.ms-excel.styles": "styles",
		"application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
		"application/vnd.openxmlformats-officedocument.custom-properties+xml": "custprops",
		"application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",
		"application/vnd.openxmlformats-officedocument.customXmlProperties+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml": "comments",
		"application/vnd.ms-excel.comments": "comments",
		"application/vnd.ms-excel.threadedcomments+xml": "threadedcomments",
		"application/vnd.ms-excel.person+xml": "people",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml": "metadata",
		"application/vnd.ms-excel.sheetMetadata": "metadata",
		"application/vnd.ms-excel.pivotTable": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.drawingml.chart+xml": "TODO",
		"application/vnd.ms-office.chartcolorstyle+xml": "TODO",
		"application/vnd.ms-office.chartstyle+xml": "TODO",
		"application/vnd.ms-office.chartex+xml": "TODO",
		"application/vnd.ms-excel.calcChain": "calcchains",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings": "TODO",
		"application/vnd.ms-office.activeX": "TODO",
		"application/vnd.ms-office.activeX+xml": "TODO",
		"application/vnd.ms-excel.attachedToolbars": "TODO",
		"application/vnd.ms-excel.connections": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml": "TODO",
		"application/vnd.ms-excel.externalLink": "links",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml": "links",
		"application/vnd.ms-excel.pivotCacheDefinition": "TODO",
		"application/vnd.ms-excel.pivotCacheRecords": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml": "TODO",
		"application/vnd.ms-excel.queryTable": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml": "TODO",
		"application/vnd.ms-excel.userNames": "TODO",
		"application/vnd.ms-excel.revisionHeaders": "TODO",
		"application/vnd.ms-excel.revisionLog": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml": "TODO",
		"application/vnd.ms-excel.tableSingleCells": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml": "TODO",
		"application/vnd.ms-excel.slicer": "TODO",
		"application/vnd.ms-excel.slicerCache": "TODO",
		"application/vnd.ms-excel.slicer+xml": "TODO",
		"application/vnd.ms-excel.slicerCache+xml": "TODO",
		"application/vnd.ms-excel.wsSortMap": "TODO",
		"application/vnd.ms-excel.table": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.theme+xml": "themes",
		"application/vnd.openxmlformats-officedocument.themeOverride+xml": "TODO",
		"application/vnd.ms-excel.Timeline+xml": "TODO",
		"application/vnd.ms-excel.TimelineCache+xml": "TODO",
		"application/vnd.ms-office.vbaProject": "vba",
		"application/vnd.ms-office.vbaProjectSignature": "TODO",
		"application/vnd.ms-office.volatileDependencies": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml": "TODO",
		"application/vnd.ms-excel.controlproperties+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.model+data": "TODO",
		"application/vnd.ms-excel.Survey+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.drawing+xml": "drawings",
		"application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.vmlDrawing": "TODO",
		"application/vnd.openxmlformats-package.relationships+xml": "rels",
		"application/vnd.openxmlformats-officedocument.oleObject": "TODO",
		"image/png": "TODO",
		sheet: "js",
	};
	function _n() {
		return {
			workbooks: [],
			sheets: [],
			charts: [],
			dialogs: [],
			macros: [],
			rels: [],
			strs: [],
			comments: [],
			threadedcomments: [],
			links: [],
			coreprops: [],
			extprops: [],
			custprops: [],
			themes: [],
			styles: [],
			calcchains: [],
			vba: [],
			drawings: [],
			metadata: [],
			people: [],
			TODO: [],
			xmlns: "",
		}
	}
	function An(e) {
		var r = _n();
		if (!e || !e.match) return r;
		var t = {}; (e.match(qr) || []).forEach(function(e) {
			var a = rt(e);
			switch (a[0].replace(Qr, "<")) {
			case "<?xml":
				break;
			case "<Types":
				r.xmlns = a["xmlns" + (a[0].match(/<(\w+):/) || ["", ""])[1]];
				break;
			case "<Default":
				t[a.Extension.toLowerCase()] = a.ContentType;
				break;
			case "<Override":
				if (r[Sn[a.ContentType]] !== undefined) r[Sn[a.ContentType]].push(a.PartName);
				break
			}
		});
		if (r.xmlns !== Rt.CT) throw new Error("Unknown Namespace: " + r.xmlns);
		r.calcchain = r.calcchains.length > 0 ? r.calcchains[0] : "";
		r.sst = r.strs.length > 0 ? r.strs[0] : "";
		r.style = r.styles.length > 0 ? r.styles[0] : "";
		r.defaults = t;
		delete r.calcchains;
		return r
	}
	var En = {
		WB: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		SHEET: "http://sheetjs.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		HLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
		VML: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
		XPATH: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
		XMISS: "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing",
		XLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
		CXML: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
		CXMLP: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
		CMNT: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
		CORE_PROPS: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
		EXT_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
		CUST_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
		SST: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
		STY: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		THEME: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
		CHART: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
		CHARTEX: "http://schemas.microsoft.com/office/2014/relationships/chartEx",
		CS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
		WS: ["http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet", ],
		DS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
		MS: "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
		IMG: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
		DRAW: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
		XLMETA: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
		TCMNT: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
		PEOPLE: "http://schemas.microsoft.com/office/2017/10/relationships/person",
		CONN: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
		VBA: "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
	};
	function Fn(e) {
		var r = e.lastIndexOf("/");
		return e.slice(0, r + 1) + "_rels/" + e.slice(r + 1) + ".rels"
	}
	function Dn(e, r) {
		var t = {
			"!id": {},
		};
		if (!e) return t;
		if (r.charAt(0) !== "/") {
			r = "/" + r
		}
		var a = {}; (e.match(qr) || []).forEach(function(e) {
			var n = rt(e);
			if (n[0] === "<Relationship") {
				var i = {};
				i.Type = n.Type;
				i.Target = it(n.Target);
				i.Id = n.Id;
				if (n.TargetMode) i.TargetMode = n.TargetMode;
				var s = n.TargetMode === "External" ? n.Target: Gr(n.Target, r);
				t[s] = i;
				a[n.Id] = i
			}
		});
		t["!id"] = a;
		return t
	}
	var Nn = "application/vnd.oasis.opendocument.spreadsheet";
	function In(e, r) {
		var t = It(e);
		var a;
		var n;
		while ((a = Pt.exec(t))) switch (a[3]) {
		case "manifest":
			break;
		case "file-entry":
			n = rt(a[0], false);
			if (n.path == "/" && n.type !== Nn) throw new Error("This OpenDocument is not a spreadsheet");
			break;
		case "encryption-data":
		case "algorithm":
		case "start-key-generation":
		case "key-derivation":
			throw new Error("Unsupported ODS Encryption");
		default:
			if (r && r.WTF) throw a;
		}
	}
	var zn = [["cp:category", "Category"], ["cp:contentStatus", "ContentStatus"], ["cp:keywords", "Keywords"], ["cp:lastModifiedBy", "LastAuthor"], ["cp:lastPrinted", "LastPrinted"], ["cp:revision", "RevNumber"], ["cp:version", "Version"], ["dc:creator", "Author"], ["dc:description", "Comments"], ["dc:identifier", "Identifier"], ["dc:language", "Language"], ["dc:subject", "Subject"], ["dc:title", "Title"], ["dcterms:created", "CreatedDate", "date"], ["dcterms:modified", "ModifiedDate", "date"], ];
	var $n = (function() {
		var e = new Array(zn.length);
		for (var r = 0; r < zn.length; ++r) {
			var t = zn[r];
			var a = "(?:" + t[0].slice(0, t[0].indexOf(":")) + ":)" + t[0].slice(t[0].indexOf(":") + 1);
			e[r] = new RegExp("<" + a + "[^>]*>([\\s\\S]*?)</" + a + ">")
		}
		return e
	})();
	function Wn(e) {
		var r = {};
		e = kt(e);
		for (var t = 0; t < zn.length; ++t) {
			var a = zn[t],
			n = e.match($n[t]);
			if (n != null && n.length > 0) r[a[1]] = it(n[1]);
			if (a[2] === "date" && r[a[1]]) r[a[1]] = wr(r[a[1]])
		}
		return r
	}
	var Vn = [["Application", "Application", "string"], ["AppVersion", "AppVersion", "string"], ["Company", "Company", "string"], ["DocSecurity", "DocSecurity", "string"], ["Manager", "Manager", "string"], ["HyperlinksChanged", "HyperlinksChanged", "bool"], ["SharedDoc", "SharedDoc", "bool"], ["LinksUpToDate", "LinksUpToDate", "bool"], ["ScaleCrop", "ScaleCrop", "bool"], ["HeadingPairs", "HeadingPairs", "raw"], ["TitlesOfParts", "TitlesOfParts", "raw"], ];
	function Gn(e, r, t, a) {
		var n = [];
		if (typeof e == "string") n = Tt(e, a);
		else for (var i = 0; i < e.length; ++i) n = n.concat(e[i].map(function(e) {
			return {
				v: e,
			}
		}));
		var s = typeof r == "string" ? Tt(r, a).map(function(e) {
			return e.v
		}) : r;
		var l = 0,
		o = 0;
		if (s.length > 0) for (var c = 0; c !== n.length; c += 2) {
			o = +n[c + 1].v;
			switch (n[c].v) {
			case "Worksheets":
			case "工作表":
			case "Листы":
			case "أوراق العمل":
			case "ワークシート":
			case "גליונות עבודה":
			case "Arbeitsblätter":
			case "Çalışma Sayfaları":
			case "Feuilles de calcul":
			case "Fogli di lavoro":
			case "Folhas de cálculo":
			case "Planilhas":
			case "Regneark":
			case "Hojas de cálculo":
			case "Werkbladen":
				t.Worksheets = o;
				t.SheetNames = s.slice(l, l + o);
				break;
			case "Named Ranges":
			case "Rangos con nombre":
			case "名前付き一覧":
			case "Benannte Bereiche":
			case "Navngivne områder":
				t.NamedRanges = o;
				t.DefinedNames = s.slice(l, l + o);
				break;
			case "Charts":
			case "Diagramme":
				t.Chartsheets = o;
				t.ChartNames = s.slice(l, l + o);
				break
			}
			l += o
		}
	}
	function Yn(e, r, t) {
		var a = {};
		if (!r) r = {};
		e = kt(e);
		Vn.forEach(function(t) {
			var n = (e.match(xt(t[0])) || [])[1];
			switch (t[2]) {
			case "string":
				if (n) r[t[1]] = it(n);
				break;
			case "bool":
				r[t[1]] = n === "true";
				break;
			case "raw":
				var i = e.match(new RegExp("<" + t[0] + "[^>]*>([\\s\\S]*?)</" + t[0] + ">"));
				if (i && i.length > 0) a[t[1]] = i[1];
				break
			}
		});
		if (a.HeadingPairs && a.TitlesOfParts) Gn(a.HeadingPairs, a.TitlesOfParts, r, t);
		return r
	}
	var Zn = /<[^>]+>[^<]*/g;
	function Kn(e, r) {
		var t = {},
		a = "";
		var n = e.match(Zn);
		if (n) for (var i = 0; i != n.length; ++i) {
			var s = n[i],
			l = rt(s);
			switch (tt(l[0])) {
			case "<?xml":
				break;
			case "<Properties":
				break;
			case "<property":
				a = it(l.name);
				break;
			case "</property>":
				a = null;
				break;
			default:
				if (s.indexOf("<vt:") === 0) {
					var o = s.split(">");
					var c = o[0].slice(4),
					f = o[1];
					switch (c) {
					case "lpstr":
					case "bstr":
					case "lpwstr":
						t[a] = it(f);
						break;
					case "bool":
						t[a] = mt(f);
						break;
					case "i1":
					case "i2":
					case "i4":
					case "i8":
					case "int":
					case "uint":
						t[a] = parseInt(f, 10);
						break;
					case "r4":
					case "r8":
					case "decimal":
						t[a] = parseFloat(f);
						break;
					case "filetime":
					case "date":
						t[a] = wr(f);
						break;
					case "cy":
					case "error":
						t[a] = it(f);
						break;
					default:
						if (c.slice(-1) == "/") break;
						if (r.WTF && typeof console !== "undefined") console.warn("Unexpected", s, c, o)
					}
				} else if (s.slice(0, 2) === "</") {} else if (r.WTF) throw new Error(s);
			}
		}
		return t
	}
	var Qn = [2, 3, 48, 49, 131, 139, 140, 245];
	var ei = (function() {
		var e = {
			1 : 437,
			2 : 850,
			3 : 1252,
			4 : 1e4,
			100 : 852,
			101 : 866,
			102 : 865,
			103 : 861,
			104 : 895,
			105 : 620,
			106 : 737,
			107 : 857,
			120 : 950,
			121 : 949,
			122 : 936,
			123 : 932,
			124 : 874,
			125 : 1255,
			126 : 1256,
			150 : 10007,
			151 : 10029,
			152 : 10006,
			200 : 1250,
			201 : 1251,
			202 : 1254,
			203 : 1253,
			0 : 20127,
			8 : 865,
			9 : 437,
			10 : 850,
			11 : 437,
			13 : 437,
			14 : 850,
			15 : 437,
			16 : 850,
			17 : 437,
			18 : 850,
			19 : 932,
			20 : 850,
			21 : 437,
			22 : 850,
			23 : 865,
			24 : 437,
			25 : 437,
			26 : 850,
			27 : 437,
			28 : 863,
			29 : 850,
			31 : 852,
			34 : 852,
			35 : 852,
			36 : 860,
			37 : 850,
			38 : 866,
			55 : 850,
			64 : 852,
			77 : 936,
			78 : 949,
			79 : 950,
			80 : 874,
			87 : 1252,
			88 : 1252,
			89 : 1252,
			108 : 863,
			134 : 737,
			135 : 852,
			136 : 857,
			204 : 1257,
			255 : 16969,
		};
		var n = lr({
			1 : 437,
			2 : 850,
			3 : 1252,
			4 : 1e4,
			100 : 852,
			101 : 866,
			102 : 865,
			103 : 861,
			104 : 895,
			105 : 620,
			106 : 737,
			107 : 857,
			120 : 950,
			121 : 949,
			122 : 936,
			123 : 932,
			124 : 874,
			125 : 1255,
			126 : 1256,
			150 : 10007,
			151 : 10029,
			152 : 10006,
			200 : 1250,
			201 : 1251,
			202 : 1254,
			203 : 1253,
			0 : 20127,
		});
		function i(r, t) {
			var n = [];
			var i = E(1);
			switch (t.type) {
			case "base64":
				i = D(C(r));
				break;
			case "binary":
				i = D(r);
				break;
			case "buffer":
			case "array":
				i = r;
				break
			}
			ya(i, 0);
			var s = i._R(1);
			var l = !!(s & 136);
			var o = false,
			c = false;
			switch (s) {
			case 2:
				break;
			case 3:
				break;
			case 48:
				o = true;
				l = true;
				break;
			case 49:
				o = true;
				l = true;
				break;
			case 131:
				break;
			case 139:
				break;
			case 140:
				c = true;
				break;
			case 245:
				break;
			default:
				throw new Error("DBF Unsupported Version: " + s.toString(16));
			}
			var f = 0,
			u = 521;
			if (s == 2) f = i._R(2);
			i.l += 3;
			if (s != 2) f = i._R(4);
			if (f > 1048576) f = 1e6;
			if (s != 2) u = i._R(2);
			var h = i._R(2);
			var d = t.codepage || 1252;
			if (s != 2) {
				i.l += 16;
				i._R(1);
				if (i[i.l] !== 0) d = e[i[i.l]];
				i.l += 1;
				i.l += 2
			}
			if (c) i.l += 36;
			var p = [],
			m = {};
			var v = Math.min(i.length, s == 2 ? 521 : u - 10 - (o ? 264 : 0));
			var g = c ? 32 : 11;
			while (i.l < v && i[i.l] != 13) {
				m = {};
				m.name = (typeof a !== "undefined" ? a.utils.decode(d, i.slice(i.l, i.l + g)) : M(i.slice(i.l, i.l + g))).replace(/[\u0000\r\n].*$/g, "");
				i.l += g;
				m.type = String.fromCharCode(i._R(1));
				if (s != 2 && !c) m.offset = i._R(4);
				m.len = i._R(1);
				if (s == 2) m.offset = i._R(2);
				m.dec = i._R(1);
				if (m.name.length) p.push(m);
				if (s != 2) i.l += c ? 13 : 14;
				switch (m.type) {
				case "B":
					if ((!o || m.len != 8) && t.WTF) console.log("Skipping " + m.name + ":" + m.type);
					break;
				case "G":
				case "P":
					if (t.WTF) console.log("Skipping " + m.name + ":" + m.type);
					break;
				case "+":
				case "0":
				case "@":
				case "C":
				case "D":
				case "F":
				case "I":
				case "L":
				case "M":
				case "N":
				case "O":
				case "T":
				case "Y":
					break;
				default:
					throw new Error("Unknown Field Type: " + m.type);
				}
			}
			if (i[i.l] !== 13) i.l = u - 1;
			if (i._R(1) !== 13) throw new Error("DBF Terminator not found " + i.l + " " + i[i.l]);
			i.l = u;
			var b = 0,
			w = 0;
			n[0] = [];
			for (w = 0; w != p.length; ++w) n[0][w] = p[w].name;
			while (f-->0) {
				if (i[i.l] === 42) {
					i.l += h;
					continue
				}++i.l;
				n[++b] = [];
				w = 0;
				for (w = 0; w != p.length; ++w) {
					var k = i.slice(i.l, i.l + p[w].len);
					i.l += p[w].len;
					ya(k, 0);
					var y = typeof a !== "undefined" ? a.utils.decode(d, k) : M(k);
					switch (p[w].type) {
					case "C":
						if (y.trim().length) n[b][w] = y.replace(/\s+$/, "");
						break;
					case "D":
						if (y.length === 8) {
							n[b][w] = new Date(Date.UTC( + y.slice(0, 4), +y.slice(4, 6) - 1, +y.slice(6, 8), 0, 0, 0, 0));
							if (! (t && t.UTC)) {
								n[b][w] = Nr(n[b][w])
							}
						} else n[b][w] = y;
						break;
					case "F":
						n[b][w] = parseFloat(y.trim());
						break;
					case "+":
					case "I":
						n[b][w] = c ? k._R(-4, "i") ^ 2147483648 : k._R(4, "i");
						break;
					case "L":
						switch (y.trim().toUpperCase()) {
						case "Y":
						case "T":
							n[b][w] = true;
							break;
						case "N":
						case "F":
							n[b][w] = false;
							break;
						case "":
						case "\0":
						case "?":
							break;
						default:
							throw new Error("DBF Unrecognized L:|" + y + "|");
						}
						break;
					case "M":
						if (!l) throw new Error("DBF Unexpected MEMO for type " + s.toString(16));
						n[b][w] = "##MEMO##" + (c ? parseInt(y.trim(), 10) : k._R(4));
						break;
					case "N":
						y = y.replace(/\u0000/g, "").trim();
						if (y && y != ".") n[b][w] = +y || 0;
						break;
					case "@":
						n[b][w] = new Date(k._R(-8, "f") - 621356832e5);
						break;
					case "T":
						{
							var x = k._R(4),
							S = k._R(4);
							if (x == 0 && S == 0) break;
							n[b][w] = new Date((x - 2440588) * 864e5 + S);
							if (! (t && t.UTC)) n[b][w] = Nr(n[b][w])
						}
						break;
					case "Y":
						n[b][w] = k._R(4, "i") / 1e4 + (k._R(4, "i") / 1e4) * Math.pow(2, 32);
						break;
					case "O":
						n[b][w] = -k._R(-8, "f");
						break;
					case "B":
						if (o && p[w].len == 8) {
							n[b][w] = k._R(8, "f");
							break
						}
					case "G":
					case "P":
						k.l += p[w].len;
						break;
					case "0":
						if (p[w].name === "_NullFlags") break;
					default:
						throw new Error("DBF Unsupported data type " + p[w].type);
					}
				}
			}
			if (s != 2) if (i.l < i.length && i[i.l++] != 26) throw new Error("DBF EOF Marker missing " + (i.l - 1) + " of " + i.length + " " + i[i.l - 1].toString(16));
			if (t && t.sheetRows) n = n.slice(0, t.sheetRows);
			t.DBF = p;
			return n
		}
		function s(e, r) {
			var t = r || {};
			if (!t.dateNF) t.dateNF = "yyyymmdd";
			var a = qa(i(e, t), t);
			a["!cols"] = t.DBF.map(function(e) {
				return {
					wch: e.len,
					DBF: e,
				}
			});
			delete t.DBF;
			return a
		}
		function l(e, r) {
			try {
				var t = Ja(s(e, r), r);
				t.bookType = "dbf";
				return t
			} catch(a) {
				if (r && r.WTF) throw a;
			}
			return {
				SheetNames: [],
				Sheets: {},
			}
		}
		var c = {
			B: 8,
			C: 250,
			L: 1,
			D: 8,
			"?": 0,
			"": 0,
		};
		function f(i, s) {
			if (!i["!ref"]) throw new Error("Cannot export empty sheet to DBF");
			var l = s || {};
			var f = r;
			if ( + l.codepage >= 0) o( + l.codepage);
			if (l.type == "string") throw new Error("Cannot write DBF to JS string");
			var u = _a();
			var h = dc(i, {
				header: 1,
				raw: true,
				cellDates: true,
			});
			var d = h[0],
			p = h.slice(1),
			m = i["!cols"] || [];
			var v = 0,
			g = 0,
			b = 0,
			w = 1;
			for (v = 0; v < d.length; ++v) {
				if (((m[v] || {}).DBF || {}).name) {
					d[v] = m[v].DBF.name; ++b;
					continue
				}
				if (d[v] == null) continue; ++b;
				if (typeof d[v] === "number") d[v] = d[v].toString(10);
				if (typeof d[v] !== "string") throw new Error("DBF Invalid column name " + d[v] + " |" + typeof d[v] + "|");
				if (d.indexOf(d[v]) !== v) for (g = 0; g < 1024; ++g) if (d.indexOf(d[v] + "_" + g) == -1) {
					d[v] += "_" + g;
					break
				}
			}
			var k = Xa(i["!ref"]);
			var y = [];
			var x = [];
			var S = [];
			for (v = 0; v <= k.e.c - k.s.c; ++v) {
				var C = "",
				_ = "",
				A = 0;
				var T = [];
				for (g = 0; g < p.length; ++g) {
					if (p[g][v] != null) T.push(p[g][v])
				}
				if (T.length == 0 || d[v] == null) {
					y[v] = "?";
					continue
				}
				for (g = 0; g < T.length; ++g) {
					switch (typeof T[g]) {
					case "number":
						_ = "B";
						break;
					case "string":
						_ = "C";
						break;
					case "boolean":
						_ = "L";
						break;
					case "object":
						_ = T[g] instanceof Date ? "D": "C";
						break;
					default:
						_ = "C"
					}
					A = Math.max(A, (typeof a !== "undefined" && typeof T[g] == "string" ? a.utils.encode(t, T[g]) : String(T[g])).length);
					C = C && C != _ ? "C": _
				}
				if (A > 250) A = 250;
				_ = ((m[v] || {}).DBF || {}).type;
				if (_ == "C") {
					if (m[v].DBF.len > A) A = m[v].DBF.len
				}
				if (C == "B" && _ == "N") {
					C = "N";
					S[v] = m[v].DBF.dec;
					A = m[v].DBF.len
				}
				x[v] = C == "C" || _ == "N" ? A: c[C] || 0;
				w += x[v];
				y[v] = C
			}
			var E = u.next(32);
			E._W(4, 318902576);
			E._W(4, p.length);
			E._W(2, 296 + 32 * b);
			E._W(2, w);
			for (v = 0; v < 4; ++v) E._W(4, 0);
			var F = +n[r] || 3;
			E._W(4, 0 | (F << 8));
			if (e[F] != +l.codepage) {
				if (l.codepage) console.error("DBF Unsupported codepage " + r + ", using 1252");
				r = 1252
			}
			for (v = 0, g = 0; v < d.length; ++v) {
				if (d[v] == null) continue;
				var D = u.next(32);
				var O = (d[v].slice(-10) + "\0\0\0\0\0\0\0\0\0\0\0").slice(0, 11);
				D._W(1, O, "sbcs");
				D._W(1, y[v] == "?" ? "C": y[v], "sbcs");
				D._W(4, g);
				D._W(1, x[v] || c[y[v]] || 0);
				D._W(1, S[v] || 0);
				D._W(1, 2);
				D._W(4, 0);
				D._W(1, 0);
				D._W(4, 0);
				D._W(4, 0);
				g += x[v] || c[y[v]] || 0
			}
			var M = u.next(264);
			M._W(4, 13);
			for (v = 0; v < 65; ++v) M._W(4, 0);
			for (v = 0; v < p.length; ++v) {
				var N = u.next(w);
				N._W(1, 0);
				for (g = 0; g < d.length; ++g) {
					if (d[g] == null) continue;
					switch (y[g]) {
					case "L":
						N._W(1, p[v][g] == null ? 63 : p[v][g] ? 84 : 70);
						break;
					case "B":
						N._W(8, p[v][g] || 0, "f");
						break;
					case "N":
						var I = "0";
						if (typeof p[v][g] == "number") I = p[v][g].toFixed(S[g] || 0);
						if (I.length > x[g]) I = I.slice(0, x[g]);
						for (b = 0; b < x[g] - I.length; ++b) N._W(1, 32);
						N._W(1, I, "sbcs");
						break;
					case "D":
						if (!p[v][g]) N._W(8, "00000000", "sbcs");
						else {
							N._W(4, ("0000" + p[v][g].getFullYear()).slice(-4), "sbcs");
							N._W(2, ("00" + (p[v][g].getMonth() + 1)).slice(-2), "sbcs");
							N._W(2, ("00" + p[v][g].getDate()).slice(-2), "sbcs")
						}
						break;
					case "C":
						var P = N.l;
						var R = String(p[v][g] != null ? p[v][g] : "").slice(0, x[g]);
						N._W(1, R, "cpstr");
						P += x[g] - N.l;
						for (b = 0; b < P; ++b) N._W(1, 32);
						break
					}
				}
			}
			r = f;
			u.next(1)._W(1, 26);
			return u.end()
		}
		return {
			to_workbook: l,
			to_sheet: s,
			from_sheet: f,
		}
	})();
	var ri = (function() {
		var e = {
			AA: "À",
			BA: "Á",
			CA: "Â",
			DA: 195,
			HA: "Ä",
			JA: 197,
			AE: "È",
			BE: "É",
			CE: "Ê",
			HE: "Ë",
			AI: "Ì",
			BI: "Í",
			CI: "Î",
			HI: "Ï",
			AO: "Ò",
			BO: "Ó",
			CO: "Ô",
			DO: 213,
			HO: "Ö",
			AU: "Ù",
			BU: "Ú",
			CU: "Û",
			HU: "Ü",
			Aa: "à",
			Ba: "á",
			Ca: "â",
			Da: 227,
			Ha: "ä",
			Ja: 229,
			Ae: "è",
			Be: "é",
			Ce: "ê",
			He: "ë",
			Ai: "ì",
			Bi: "í",
			Ci: "î",
			Hi: "ï",
			Ao: "ò",
			Bo: "ó",
			Co: "ô",
			Do: 245,
			Ho: "ö",
			Au: "ù",
			Bu: "ú",
			Cu: "û",
			Hu: "ü",
			KC: "Ç",
			Kc: "ç",
			q: "æ",
			z: "œ",
			a: "Æ",
			j: "Œ",
			DN: 209,
			Dn: 241,
			Hy: 255,
			S: 169,
			c: 170,
			R: 174,
			"B ": 180,
			0 : 176,
			1 : 177,
			2 : 178,
			3 : 179,
			5 : 181,
			6 : 182,
			7 : 183,
			Q: 185,
			k: 186,
			b: 208,
			i: 216,
			l: 222,
			s: 240,
			y: 248,
			"!": 161,
			'"': 162,
			"#": 163,
			"(": 164,
			"%": 165,
			"'": 167,
			"H ": 168,
			"+": 171,
			";": 187,
			"<": 188,
			"=": 189,
			">": 190,
			"?": 191,
			"{": 223,
		};
		var r = new RegExp("N(" + ir(e).join("|").replace(/\|\|\|/, "|\\||").replace(/([?()+])/g, "\\$1").replace("{", "\\{") + "|\\|)", "gm");
		try {
			r = new RegExp("N(" + ir(e).join("|").replace(/\|\|\|/, "|\\||").replace(/([?()+])/g, "\\$1") + "|\\|)", "gm")
		} catch(t) {}
		var n = function(r, t) {
			var a = e[t];
			return typeof a == "number" ? v(a) : a
		};
		var i = function(e, r, t) {
			var a = ((r.charCodeAt(0) - 32) << 4) | (t.charCodeAt(0) - 48);
			return a == 59 ? e: v(a)
		};
		e["|"] = 254;
		var s = function(e) {
			return e.replace(/\n/g, " :").replace(/\r/g, " =")
		};
		function l(e, r) {
			switch (r.type) {
			case "base64":
				return c(C(e), r);
			case "binary":
				return c(e, r);
			case "buffer":
				return c(_ && Buffer.isBuffer(e) ? e.toString("binary") : M(e), r);
			case "array":
				return c(kr(e), r)
			}
			throw new Error("Unrecognized type " + r.type);
		}
		function c(e, t) {
			var s = e.split(/[\n\r]+/),
			l = -1,
			c = -1,
			f = 0,
			u = 0,
			h = [];
			var d = [];
			var p = null;
			var m = {},
			v = [],
			g = [],
			b = [];
			var w = 0,
			k;
			var y = {
				Workbook: {
					WBProps: {},
					Names: [],
				},
			};
			if ( + t.codepage >= 0) o( + t.codepage);
			for (; f !== s.length; ++f) {
				w = 0;
				var x = s[f].trim().replace(/\x1B([\x20-\x2F])([\x30-\x3F])/g, i).replace(r, n);
				var S = x.replace(/;;/g, "\0").split(";").map(function(e) {
					return e.replace(/\u0000/g, ";")
				});
				var C = S[0],
				_;
				if (x.length > 0) switch (C) {
				case "ID":
					break;
				case "E":
					break;
				case "B":
					break;
				case "O":
					for (u = 1; u < S.length; ++u) switch (S[u].charAt(0)) {
					case "V":
						{
							var A = parseInt(S[u].slice(1), 10);
							if (A >= 1 && A <= 4) y.Workbook.WBProps.date1904 = true
						}
						break
					}
					break;
				case "W":
					break;
				case "P":
					switch (S[1].charAt(0)) {
					case "P":
						d.push(x.slice(3).replace(/;;/g, ";"));
						break
					}
					break;
				case "NN":
					{
						var T = {
							Sheet: 0,
						};
						for (u = 1; u < S.length; ++u) switch (S[u].charAt(0)) {
						case "N":
							T.Name = S[u].slice(1);
							break;
						case "E":
							T.Ref = ((t && t.sheet) || "Sheet1") + "!" + Os(S[u].slice(1));
							break
						}
						y.Workbook.Names.push(T)
					}
					break;
				case "C":
					var E = false,
					F = false,
					D = false,
					O = false,
					M = -1,
					N = -1,
					I = "",
					P = "z";
					var R = "";
					for (u = 1; u < S.length; ++u) switch (S[u].charAt(0)) {
					case "A":
						R = S[u].slice(1);
						break;
					case "X":
						c = parseInt(S[u].slice(1), 10) - 1;
						F = true;
						break;
					case "Y":
						l = parseInt(S[u].slice(1), 10) - 1;
						if (!F) c = 0;
						for (k = h.length; k <= l; ++k) h[k] = [];
						break;
					case "K":
						_ = S[u].slice(1);
						if (_.charAt(0) === '"') {
							_ = _.slice(1, _.length - 1);
							P = "s"
						} else if (_ === "TRUE" || _ === "FALSE") {
							_ = _ === "TRUE";
							P = "b"
						} else if (!isNaN(Sr(_))) {
							_ = Sr(_);
							P = "n";
							if (p !== null && Re(p) && t.cellDates) {
								_ = pr(y.Workbook.WBProps.date1904 ? _ + 1462 : _);
								P = typeof _ == "number" ? "n": "d"
							}
						}
						if (typeof a !== "undefined" && typeof _ == "string" && (t || {}).type != "string" && (t || {}).codepage) _ = a.utils.decode(t.codepage, _);
						E = true;
						break;
					case "E":
						O = true;
						I = Os(S[u].slice(1), {
							r: l,
							c: c,
						});
						break;
					case "S":
						D = true;
						break;
					case "G":
						break;
					case "R":
						M = parseInt(S[u].slice(1), 10) - 1;
						break;
					case "C":
						N = parseInt(S[u].slice(1), 10) - 1;
						break;
					default:
						if (t && t.WTF) throw new Error("SYLK bad record " + x);
					}
					if (E) {
						if (!h[l][c]) h[l][c] = {
							t: P,
							v: _,
						};
						else {
							h[l][c].t = P;
							h[l][c].v = _
						}
						if (p) h[l][c].z = p;
						if (t.cellText !== false && p) h[l][c].w = $e(h[l][c].z, h[l][c].v, {
							date1904: y.Workbook.WBProps.date1904,
						});
						p = null
					}
					if (D) {
						if (O) throw new Error("SYLK shared formula cannot have own formula");
						var L = M > -1 && h[M][N];
						if (!L || !L[1]) throw new Error("SYLK shared formula cannot find base");
						I = Ps(L[1], {
							r: l - M,
							c: c - N,
						})
					}
					if (I) {
						if (!h[l][c]) h[l][c] = {
							t: "n",
							f: I,
						};
						else h[l][c].f = I
					}
					if (R) {
						if (!h[l][c]) h[l][c] = {
							t: "z",
						};
						h[l][c].c = [{
							a: "SheetJSYLK",
							t: R,
						},
						]
					}
					break;
				case "F":
					var B = 0;
					for (u = 1; u < S.length; ++u) switch (S[u].charAt(0)) {
					case "X":
						c = parseInt(S[u].slice(1), 10) - 1; ++B;
						break;
					case "Y":
						l = parseInt(S[u].slice(1), 10) - 1;
						for (k = h.length; k <= l; ++k) h[k] = [];
						break;
					case "M":
						w = parseInt(S[u].slice(1), 10) / 20;
						break;
					case "F":
						break;
					case "G":
						break;
					case "P":
						p = d[parseInt(S[u].slice(1), 10)];
						break;
					case "S":
						break;
					case "D":
						break;
					case "N":
						break;
					case "W":
						b = S[u].slice(1).split(" ");
						for (k = parseInt(b[0], 10); k <= parseInt(b[1], 10); ++k) {
							w = parseInt(b[2], 10);
							g[k - 1] = w === 0 ? {
								hidden: true,
							}: {
								wch: w,
							}
						}
						break;
					case "C":
						c = parseInt(S[u].slice(1), 10) - 1;
						if (!g[c]) g[c] = {};
						break;
					case "R":
						l = parseInt(S[u].slice(1), 10) - 1;
						if (!v[l]) v[l] = {};
						if (w > 0) {
							v[l].hpt = w;
							v[l].hpx = Li(w)
						} else if (w === 0) v[l].hidden = true;
						break;
					default:
						if (t && t.WTF) throw new Error("SYLK bad record " + x);
					}
					if (B < 1) p = null;
					break;
				default:
					if (t && t.WTF) throw new Error("SYLK bad record " + x);
				}
			}
			if (v.length > 0) m["!rows"] = v;
			if (g.length > 0) m["!cols"] = g;
			g.forEach(function(e) {
				Ni(e)
			});
			if (t && t.sheetRows) h = h.slice(0, t.sheetRows);
			return [h, m, y]
		}
		function f(e, r) {
			var t = l(e, r);
			var a = t[0],
			n = t[1],
			i = t[2];
			var s = yr(r);
			s.date1904 = (((i || {}).Workbook || {}).WBProps || {}).date1904;
			var o = qa(a, s);
			ir(n).forEach(function(e) {
				o[e] = n[e]
			});
			var c = Ja(o, r);
			ir(i).forEach(function(e) {
				c[e] = i[e]
			});
			c.bookType = "sylk";
			return c
		}
		function u(e, r, t, a, n, i) {
			var s = "C;Y" + (t + 1) + ";X" + (a + 1) + ";K";
			switch (e.t) {
			case "n":
				s += e.v || 0;
				if (e.f && !e.F) s += ";E" + Is(e.f, {
					r: t,
					c: a,
				});
				break;
			case "b":
				s += e.v ? "TRUE": "FALSE";
				break;
			case "e":
				s += e.w || e.v;
				break;
			case "d":
				s += dr(wr(e.v, i), i);
				break;
			case "s":
				s += '"' + (e.v == null ? "": String(e.v)).replace(/"/g, "").replace(/;/g, ";;") + '"';
				break
			}
			return s
		}
		function h(e, r, t) {
			var a = "C;Y" + (r + 1) + ";X" + (t + 1) + ";A";
			a += s(e.map(function(e) {
				return e.t
			}).join(""));
			return a
		}
		function d(e, r) {
			r.forEach(function(r, t) {
				var a = "F;W" + (t + 1) + " " + (t + 1) + " ";
				if (r.hidden) a += "0";
				else {
					if (typeof r.width == "number" && !r.wpx) r.wpx = Ei(r.width);
					if (typeof r.wpx == "number" && !r.wch) r.wch = Fi(r.wpx);
					if (typeof r.wch == "number") a += Math.round(r.wch)
				}
				if (a.charAt(a.length - 1) != " ") e.push(a)
			})
		}
		function p(e, r) {
			r.forEach(function(r, t) {
				var a = "F;";
				if (r.hidden) a += "M0;";
				else if (r.hpt) a += "M" + 20 * r.hpt + ";";
				else if (r.hpx) a += "M" + 20 * Ri(r.hpx) + ";";
				if (a.length > 2) e.push(a + "R" + (t + 1))
			})
		}
		function m(e, r, t) {
			if (!r) r = {};
			r._formats = ["General"];
			var a = ["ID;PSheetJS;N;E"],
			n = [];
			var i = Xa(e["!ref"] || "A1"),
			s;
			var l = e["!data"] != null;
			var o = "\r\n";
			var c = (((t || {}).Workbook || {}).WBProps || {}).date1904;
			var f = "General";
			a.push("P;PGeneral");
			var m = i.s.r,
			v = i.s.c,
			g = [];
			if (e["!ref"]) for (m = i.s.r; m <= i.e.r; ++m) {
				if (l && !e["!data"][m]) continue;
				g = [];
				for (v = i.s.c; v <= i.e.c; ++v) {
					s = l ? e["!data"][m][v] : e[Ra(v) + Ma(m)];
					if (!s || !s.c) continue;
					g.push(h(s.c, m, v))
				}
				if (g.length) n.push(g.join(o))
			}
			if (e["!ref"]) for (m = i.s.r; m <= i.e.r; ++m) {
				if (l && !e["!data"][m]) continue;
				g = [];
				for (v = i.s.c; v <= i.e.c; ++v) {
					s = l ? e["!data"][m][v] : e[Ra(v) + Ma(m)];
					if (!s || (s.v == null && (!s.f || s.F))) continue;
					if ((s.z || (s.t == "d" ? q[14] : "General")) != f) {
						var b = r._formats.indexOf(s.z);
						if (b == -1) {
							r._formats.push(s.z);
							b = r._formats.length - 1;
							a.push("P;P" + s.z.replace(/;/g, ";;"))
						}
						g.push("F;P" + b + ";Y" + (m + 1) + ";X" + (v + 1))
					}
					g.push(u(s, e, m, v, r, c))
				}
				n.push(g.join(o))
			}
			a.push("F;P0;DG0G8;M255");
			if (e["!cols"]) d(a, e["!cols"]);
			if (e["!rows"]) p(a, e["!rows"]);
			if (e["!ref"]) a.push("B;Y" + (i.e.r - i.s.r + 1) + ";X" + (i.e.c - i.s.c + 1) + ";D" + [i.s.c, i.s.r, i.e.c, i.e.r].join(" "));
			a.push("O;L;D;B" + (c ? ";V4": "") + ";K47;G100 0.001");
			delete r._formats;
			return a.join(o) + o + n.join(o) + o + "E" + o
		}
		return {
			to_workbook: f,
			from_sheet: m,
		}
	})();
	var ti = (function() {
		function e(e, t) {
			switch (t.type) {
			case "base64":
				return r(C(e), t);
			case "binary":
				return r(e, t);
			case "buffer":
				return r(_ && Buffer.isBuffer(e) ? e.toString("binary") : M(e), t);
			case "array":
				return r(kr(e), t)
			}
			throw new Error("Unrecognized type " + t.type);
		}
		function r(e, r) {
			var t = e.split("\n"),
			a = -1,
			n = -1,
			i = 0,
			s = [];
			for (; i !== t.length; ++i) {
				if (t[i].trim() === "BOT") {
					s[++a] = [];
					n = 0;
					continue
				}
				if (a < 0) continue;
				var l = t[i].trim().split(",");
				var o = l[0],
				c = l[1]; ++i;
				var f = t[i] || "";
				while ((f.match(/["]/g) || []).length & 1 && i < t.length - 1) f += "\n" + t[++i];
				f = f.trim();
				switch ( + o) {
				case - 1 : if (f === "BOT") {
						s[++a] = [];
						n = 0;
						continue
					} else if (f !== "EOD") throw new Error("Unrecognized DIF special command " + f);
					break;
				case 0:
					if (f === "TRUE") s[a][n] = true;
					else if (f === "FALSE") s[a][n] = false;
					else if (!isNaN(Sr(c))) s[a][n] = Sr(c);
					else if (!isNaN(Or(c).getDate())) {
						s[a][n] = wr(c);
						if (! (r && r.UTC)) {
							s[a][n] = Nr(s[a][n])
						}
					} else s[a][n] = c; ++n;
					break;
				case 1:
					f = f.slice(1, f.length - 1);
					f = f.replace(/""/g, '"');
					if (w && f && f.match(/^=".*"$/)) f = f.slice(2, -1);
					s[a][n++] = f !== "" ? f: null;
					break
				}
				if (f === "EOD") break
			}
			if (r && r.sheetRows) s = s.slice(0, r.sheetRows);
			return s
		}
		function t(r, t) {
			return qa(e(r, t), t)
		}
		function a(e, r) {
			var a = Ja(t(e, r), r);
			a.bookType = "dif";
			return a
		}
		function n(e, r) {
			return "0," + String(e) + "\r\n" + r
		}
		function i(e) {
			return '1,0\r\n"' + e.replace(/"/g, '""') + '"'
		}
		function s(e) {
			var r = w;
			if (!e["!ref"]) throw new Error("Cannot export empty sheet to DIF");
			var t = Xa(e["!ref"]);
			var a = e["!data"] != null;
			var s = ['TABLE\r\n0,1\r\n"sheetjs"\r\n', "VECTORS\r\n0," + (t.e.r - t.s.r + 1) + '\r\n""\r\n', "TUPLES\r\n0," + (t.e.c - t.s.c + 1) + '\r\n""\r\n', 'DATA\r\n0,0\r\n""\r\n', ];
			for (var l = t.s.r; l <= t.e.r; ++l) {
				var o = a ? e["!data"][l] : [];
				var c = "-1,0\r\nBOT\r\n";
				for (var f = t.s.c; f <= t.e.c; ++f) {
					var u = a ? o && o[f] : e[$a({
						r: l,
						c: f,
					})];
					if (u == null) {
						c += '1,0\r\n""\r\n';
						continue
					}
					switch (u.t) {
					case "n":
						if (r) {
							if (u.w != null) c += "0," + u.w + "\r\nV";
							else if (u.v != null) c += n(u.v, "V");
							else if (u.f != null && !u.F) c += i("=" + u.f);
							else c += '1,0\r\n""'
						} else {
							if (u.v == null) c += '1,0\r\n""';
							else c += n(u.v, "V")
						}
						break;
					case "b":
						c += u.v ? n(1, "TRUE") : n(0, "FALSE");
						break;
					case "s":
						c += i(!r || isNaN( + u.v) ? u.v: '="' + u.v + '"');
						break;
					case "d":
						if (!u.w) u.w = $e(u.z || q[14], dr(wr(u.v)));
						if (r) c += n(u.w, "V");
						else c += i(u.w);
						break;
					default:
						c += '1,0\r\n""'
					}
					c += "\r\n"
				}
				s.push(c)
			}
			return s.join("") + "-1,0\r\nEOD"
		}
		return {
			to_workbook: a,
			to_sheet: t,
			from_sheet: s,
		}
	})();
	var ai = (function() {
		function e(e) {
			return e.replace(/\\b/g, "\\").replace(/\\c/g, ":").replace(/\\n/g, "\n")
		}
		function r(e) {
			return e.replace(/\\/g, "\\b").replace(/:/g, "\\c").replace(/\n/g, "\\n")
		}
		function t(r, t) {
			var a = r.split("\n"),
			n = -1,
			i = -1,
			s = 0,
			l = [];
			for (; s !== a.length; ++s) {
				var o = a[s].trim().split(":");
				if (o[0] !== "cell") continue;
				var c = za(o[1]);
				if (l.length <= c.r) for (n = l.length; n <= c.r; ++n) if (!l[n]) l[n] = [];
				n = c.r;
				i = c.c;
				switch (o[2]) {
				case "t":
					l[n][i] = e(o[3]);
					break;
				case "v":
					l[n][i] = +o[3];
					break;
				case "vtf":
					var f = o[o.length - 1];
				case "vtc":
					switch (o[3]) {
					case "nl":
						l[n][i] = +o[4] ? true: false;
						break;
					default:
						l[n][i] = +o[4];
						break
					}
					if (o[2] == "vtf") l[n][i] = [l[n][i], f]
				}
			}
			if (t && t.sheetRows) l = l.slice(0, t.sheetRows);
			return l
		}
		function a(e, r) {
			return qa(t(e, r), r)
		}
		function n(e, r) {
			return Ja(a(e, r), r)
		}
		var i = ["socialcalc:version:1.5", "MIME-Version: 1.0", "Content-Type: multipart/mixed; boundary=SocialCalcSpreadsheetControlSave", ].join("\n");
		var s = ["--SocialCalcSpreadsheetControlSave", "Content-type: text/plain; charset=UTF-8", ].join("\n") + "\n";
		var l = ["# SocialCalc Spreadsheet Control Save", "part:sheet"].join("\n");
		var o = "--SocialCalcSpreadsheetControlSave--";
		function c(e) {
			if (!e || !e["!ref"]) return "";
			var t = [],
			a = [],
			n,
			i = "";
			var s = Wa(e["!ref"]);
			var l = e["!data"] != null;
			for (var o = s.s.r; o <= s.e.r; ++o) {
				for (var c = s.s.c; c <= s.e.c; ++c) {
					i = $a({
						r: o,
						c: c,
					});
					n = l ? (e["!data"][o] || [])[c] : e[i];
					if (!n || n.v == null || n.t === "z") continue;
					a = ["cell", i, "t"];
					switch (n.t) {
					case "s":
					case "str":
						a.push(r(n.v));
						break;
					case "n":
						if (!n.f) {
							a[2] = "v";
							a[3] = n.v
						} else {
							a[2] = "vtf";
							a[3] = "n";
							a[4] = n.v;
							a[5] = r(n.f)
						}
						break;
					case "b":
						a[2] = "vt" + (n.f ? "f": "c");
						a[3] = "nl";
						a[4] = n.v ? "1": "0";
						a[5] = r(n.f || (n.v ? "TRUE": "FALSE"));
						break;
					case "d":
						var f = dr(wr(n.v));
						a[2] = "vtc";
						a[3] = "nd";
						a[4] = "" + f;
						a[5] = n.w || $e(n.z || q[14], f);
						break;
					case "e":
						continue
					}
					t.push(a.join(":"))
				}
			}
			t.push("sheet:c:" + (s.e.c - s.s.c + 1) + ":r:" + (s.e.r - s.s.r + 1) + ":tvf:1");
			t.push("valueformat:1:text-wiki");
			return t.join("\n")
		}
		function f(e) {
			return [i, s, l, s, c(e), o].join("\n")
		}
		return {
			to_workbook: n,
			to_sheet: a,
			from_sheet: f,
		}
	})();
	var ni = (function() {
		function e(e, r, t, a, n) {
			if (n.raw) r[t][a] = e;
			else if (e === "") {} else if (e === "TRUE") r[t][a] = true;
			else if (e === "FALSE") r[t][a] = false;
			else if (!isNaN(Sr(e))) r[t][a] = Sr(e);
			else if (!isNaN(Or(e).getDate())) r[t][a] = wr(e);
			else r[t][a] = e
		}
		function r(r, t) {
			var a = t || {};
			var n = [];
			if (!r || r.length === 0) return n;
			var i = r.split(/[\r\n]/);
			var s = i.length - 1;
			while (s >= 0 && i[s].length === 0)--s;
			var l = 10,
			o = 0;
			var c = 0;
			for (; c <= s; ++c) {
				o = i[c].indexOf(" ");
				if (o == -1) o = i[c].length;
				else o++;
				l = Math.max(l, o)
			}
			for (c = 0; c <= s; ++c) {
				n[c] = [];
				var f = 0;
				e(i[c].slice(0, l).trim(), n, c, f, a);
				for (f = 1; f <= (i[c].length - l) / 10 + 1; ++f) e(i[c].slice(l + (f - 1) * 10, l + f * 10).trim(), n, c, f, a)
			}
			if (a.sheetRows) n = n.slice(0, a.sheetRows);
			return n
		}
		var t = {
			44 : ",",
			9 : "\t",
			59 : ";",
			124 : "|",
		};
		var n = {
			44 : 3,
			9 : 2,
			59 : 1,
			124 : 0,
		};
		function i(e) {
			var r = {},
			a = false,
			i = 0,
			s = 0;
			for (; i < e.length; ++i) {
				if ((s = e.charCodeAt(i)) == 34) a = !a;
				else if (!a && s in t) r[s] = (r[s] || 0) + 1
			}
			s = [];
			for (i in r) if (Object.prototype.hasOwnProperty.call(r, i)) {
				s.push([r[i], i])
			}
			if (!s.length) {
				r = n;
				for (i in r) if (Object.prototype.hasOwnProperty.call(r, i)) {
					s.push([r[i], i])
				}
			}
			s.sort(function(e, r) {
				return e[0] - r[0] || n[e[1]] - n[r[1]]
			});
			return t[s.pop()[1]] || 44
		}
		function s(e, r) {
			var t = r || {};
			var a = "";
			if (b != null && t.dense == null) t.dense = b;
			var n = {};
			if (t.dense) n["!data"] = [];
			var s = {
				s: {
					c: 0,
					r: 0,
				},
				e: {
					c: 0,
					r: 0,
				},
			};
			if (e.slice(0, 4) == "sep=") {
				if (e.charCodeAt(5) == 13 && e.charCodeAt(6) == 10) {
					a = e.charAt(4);
					e = e.slice(7)
				} else if (e.charCodeAt(5) == 13 || e.charCodeAt(5) == 10) {
					a = e.charAt(4);
					e = e.slice(6)
				} else a = i(e.slice(0, 1024))
			} else if (t && t.FS) a = t.FS;
			else a = i(e.slice(0, 1024));
			var l = 0,
			o = 0,
			c = 0;
			var f = 0,
			u = 0,
			h = a.charCodeAt(0),
			d = false,
			p = 0,
			m = e.charCodeAt(0);
			var v = t.dateNF != null ? Ye(t.dateNF) : null;
			function g() {
				var r = e.slice(f, u);
				if (r.slice(-1) == "\r") r = r.slice(0, -1);
				var a = {};
				if (r.charAt(0) == '"' && r.charAt(r.length - 1) == '"') r = r.slice(1, -1).replace(/""/g, '"');
				if (t.cellText !== false) a.w = r;
				if (r.length === 0) a.t = "z";
				else if (t.raw) {
					a.t = "s";
					a.v = r
				} else if (r.trim().length === 0) {
					a.t = "s";
					a.v = r
				} else if (r.charCodeAt(0) == 61) {
					if (r.charCodeAt(1) == 34 && r.charCodeAt(r.length - 1) == 34) {
						a.t = "s";
						a.v = r.slice(2, -1).replace(/""/g, '"')
					} else if (Ls(r)) {
						a.t = "s";
						a.f = r.slice(1);
						a.v = r
					} else {
						a.t = "s";
						a.v = r
					}
				} else if (r == "TRUE") {
					a.t = "b";
					a.v = true
				} else if (r == "FALSE") {
					a.t = "b";
					a.v = false
				} else if (!isNaN((c = Sr(r)))) {
					a.t = "n";
					a.v = c
				} else if (!isNaN((c = Or(r)).getDate()) || (v && r.match(v))) {
					a.z = t.dateNF || q[14];
					if (v && r.match(v)) {
						var i = Je(r, t.dateNF, r.match(v) || []);
						c = wr(i);
						if (t && t.UTC === false) c = Nr(c)
					} else if (t && t.UTC === false) c = Nr(c);
					else if (t.cellText !== false && t.dateNF) a.w = $e(a.z, c);
					if (t.cellDates) {
						a.t = "d";
						a.v = c
					} else {
						a.t = "n";
						a.v = dr(c)
					}
					if (!t.cellNF) delete a.z
				} else {
					a.t = "s";
					a.v = r
				}
				if (a.t == "z") {} else if (t.dense) {
					if (!n["!data"][l]) n["!data"][l] = [];
					n["!data"][l][o] = a
				} else n[$a({
					c: o,
					r: l,
				})] = a;
				f = u + 1;
				m = e.charCodeAt(f);
				if (s.e.c < o) s.e.c = o;
				if (s.e.r < l) s.e.r = l;
				if (p == h)++o;
				else {
					o = 0; ++l;
					if (t.sheetRows && t.sheetRows <= l) return true
				}
			}
			e: for (; u < e.length; ++u) switch ((p = e.charCodeAt(u))) {
			case 34:
				if (m === 34) d = !d;
				break;
			case 13:
				if (d) break;
				if (e.charCodeAt(u + 1) == 10)++u;
			case h:
			case 10:
				if (!d && g()) break e;
				break;
			default:
				break
			}
			if (u - f > 0) g();
			n["!ref"] = ja(s);
			return n
		}
		function l(e, t) {
			if (! (t && t.PRN)) return s(e, t);
			if (t.FS) return s(e, t);
			if (e.slice(0, 4) == "sep=") return s(e, t);
			if (e.indexOf("\t") >= 0 || e.indexOf(",") >= 0 || e.indexOf(";") >= 0) return s(e, t);
			return qa(r(e, t), t)
		}
		function o(e, r) {
			var t = "",
			n = r.type == "string" ? [0, 0, 0, 0] : $o(e, r);
			switch (r.type) {
			case "base64":
				t = C(e);
				break;
			case "binary":
				t = e;
				break;
			case "buffer":
				if (r.codepage == 65001) t = e.toString("utf8");
				else if (r.codepage && typeof a !== "undefined") t = a.utils.decode(r.codepage, e);
				else t = _ && Buffer.isBuffer(e) ? e.toString("binary") : M(e);
				break;
			case "array":
				t = kr(e);
				break;
			case "string":
				t = e;
				break;
			default:
				throw new Error("Unrecognized type " + r.type);
			}
			if (n[0] == 239 && n[1] == 187 && n[2] == 191) t = kt(t.slice(3));
			else if (r.type != "string" && r.type != "buffer" && r.codepage == 65001) t = kt(t);
			else if (r.type == "binary" && typeof a !== "undefined" && r.codepage) t = a.utils.decode(r.codepage, a.utils.encode(28591, t));
			if (t.slice(0, 19) == "socialcalc:version:") return ai.to_sheet(r.type == "string" ? t: kt(t), r);
			return l(t, r)
		}
		function c(e, r) {
			return Ja(o(e, r), r)
		}
		function f(e) {
			var r = [];
			if (!e["!ref"]) return "";
			var t = Xa(e["!ref"]),
			a;
			var n = e["!data"] != null;
			for (var i = t.s.r; i <= t.e.r; ++i) {
				var s = [];
				for (var l = t.s.c; l <= t.e.c; ++l) {
					var o = $a({
						r: i,
						c: l,
					});
					a = n ? (e["!data"][i] || [])[l] : e[o];
					if (!a || a.v == null) {
						s.push("          ");
						continue
					}
					var c = (a.w || (Ya(a), a.w) || "").slice(0, 10);
					while (c.length < 10) c += " ";
					s.push(c + (l === 0 ? " ": ""))
				}
				r.push(s.join(""))
			}
			return r.join("\n")
		}
		return {
			to_workbook: c,
			to_sheet: o,
			from_sheet: f,
		}
	})();
	function ii(e, r) {
		var t = r || {},
		a = !!t.WTF;
		t.WTF = true;
		try {
			var n = ri.to_workbook(e, t);
			t.WTF = a;
			return n
		} catch(i) {
			t.WTF = a;
			if (!i.message.match(/SYLK bad record ID/) && a) throw i;
			return ni.to_workbook(e, r)
		}
	}
	function si(e) {
		var r = {},
		t = e.match(qr),
		a = 0;
		var n = false;
		if (t) for (; a != t.length; ++a) {
			var s = rt(t[a]);
			switch (s[0].replace(/\w*:/g, "")) {
			case "<condense":
				break;
			case "<extend":
				break;
			case "<shadow":
				if (!s.val) break;
			case "<shadow>":
			case "<shadow/>":
				r.shadow = 1;
				break;
			case "</shadow>":
				break;
			case "<charset":
				if (s.val == "1") break;
				r.cp = i[parseInt(s.val, 10)];
				break;
			case "<outline":
				if (!s.val) break;
			case "<outline>":
			case "<outline/>":
				r.outline = 1;
				break;
			case "</outline>":
				break;
			case "<rFont":
				r.name = s.val;
				break;
			case "<sz":
				r.sz = s.val;
				break;
			case "<strike":
				if (!s.val) break;
			case "<strike>":
			case "<strike/>":
				r.strike = 1;
				break;
			case "</strike>":
				break;
			case "<u":
				if (!s.val) break;
				switch (s.val) {
				case "double":
					r.uval = "double";
					break;
				case "singleAccounting":
					r.uval = "single-accounting";
					break;
				case "doubleAccounting":
					r.uval = "double-accounting";
					break
				}
			case "<u>":
			case "<u/>":
				r.u = 1;
				break;
			case "</u>":
				break;
			case "<b":
				if (s.val == "0") break;
			case "<b>":
			case "<b/>":
				r.b = 1;
				break;
			case "</b>":
				break;
			case "<i":
				if (s.val == "0") break;
			case "<i>":
			case "<i/>":
				r.i = 1;
				break;
			case "</i>":
				break;
			case "<color":
				if (s.rgb) r.color = s.rgb.slice(2, 8);
				break;
			case "<color>":
			case "<color/>":
			case "</color>":
				break;
			case "<family":
				r.family = s.val;
				break;
			case "<family>":
			case "<family/>":
			case "</family>":
				break;
			case "<vertAlign":
				r.valign = s.val;
				break;
			case "<vertAlign>":
			case "<vertAlign/>":
			case "</vertAlign>":
				break;
			case "<scheme":
				break;
			case "<scheme>":
			case "<scheme/>":
			case "</scheme>":
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
				break;
			case "<ext":
				n = true;
				break;
			case "</ext>":
				n = false;
				break;
			default:
				if (s[0].charCodeAt(1) !== 47 && !n) throw new Error("Unrecognized rich format " + s[0]);
			}
		}
		return r
	}
	var li = (function() {
		var e = xt("t"),
		r = xt("rPr");
		function t(t) {
			var a = t.match(e);
			if (!a) return {
				t: "s",
				v: "",
			};
			var n = {
				t: "s",
				v: it(a[1]),
			};
			var i = t.match(r);
			if (i) n.s = si(i[1]);
			return n
		}
		var a = /<(?:\w+:)?r>/g,
		n = /<\/(?:\w+:)?r>/;
		return function i(e) {
			return e.replace(a, "").split(n).map(t).filter(function(e) {
				return e.v
			})
		}
	})();
	var oi = (function Yc() {
		var e = /(\r\n|\n)/g;
		function r(e, r, t) {
			var a = [];
			if (e.u) a.push("text-decoration: underline;");
			if (e.uval) a.push("text-underline-style:" + e.uval + ";");
			if (e.sz) a.push("font-size:" + e.sz + "pt;");
			if (e.outline) a.push("text-effect: outline;");
			if (e.shadow) a.push("text-shadow: auto;");
			r.push('<span style="' + a.join("") + '">');
			if (e.b) {
				r.push("<b>");
				t.push("</b>")
			}
			if (e.i) {
				r.push("<i>");
				t.push("</i>")
			}
			if (e.strike) {
				r.push("<s>");
				t.push("</s>")
			}
			var n = e.valign || "";
			if (n == "superscript" || n == "super") n = "sup";
			else if (n == "subscript") n = "sub";
			if (n != "") {
				r.push("<" + n + ">");
				t.push("</" + n + ">")
			}
			t.push("</span>");
			return e
		}
		function t(t) {
			var a = [[], t.v, []];
			if (!t.v) return "";
			if (t.s) r(t.s, a[0], a[2]);
			return a[0].join("") + a[1].replace(e, "<br/>") + a[2].join("")
		}
		return function a(e) {
			return e.map(t).join("")
		}
	})();
	var ci = /<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/g,
	fi = /<(?:\w+:)?r\b[^>]*>/;
	var ui = /<(?:\w+:)?rPh.*?>([\s\S]*?)<\/(?:\w+:)?rPh>/g;
	function hi(e, r) {
		var t = r ? r.cellHTML: true;
		var a = {};
		if (!e) return {
			t: "",
		};
		if (e.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
			a.t = it(kt(e.slice(e.indexOf(">") + 1).split(/<\/(?:\w+:)?t>/)[0] || ""), true);
			a.r = kt(e);
			if (t) a.h = ut(a.t)
		} else if (e.match(fi)) {
			a.r = kt(e);
			a.t = it(kt((e.replace(ui, "").match(ci) || []).join("").replace(qr, "")), true);
			if (t) a.h = oi(li(a.r))
		}
		return a
	}
	var di = /<(?:\w+:)?sst([^>]*)>([\s\S]*)<\/(?:\w+:)?sst>/;
	var pi = /<(?:\w+:)?(?:si|sstItem)>/g;
	var mi = /<\/(?:\w+:)?(?:si|sstItem)>/;
	function vi(e, r) {
		var t = [],
		a = "";
		if (!e) return t;
		var n = e.match(di);
		if (n) {
			a = n[2].replace(pi, "").split(mi);
			for (var i = 0; i != a.length; ++i) {
				var s = hi(a[i].trim(), r);
				if (s != null) t[t.length] = s
			}
			n = rt(n[1]);
			t.Count = n.count;
			t.Unique = n.uniqueCount
		}
		return t
	}
	function wi(e) {
		var r = e.slice(e[0] === "#" ? 1 : 0).slice(0, 6);
		return [parseInt(r.slice(0, 2), 16), parseInt(r.slice(2, 4), 16), parseInt(r.slice(4, 6), 16), ]
	}
	function ki(e) {
		for (var r = 0,
		t = 1; r != 3; ++r) t = t * 256 + (e[r] > 255 ? 255 : e[r] < 0 ? 0 : e[r]);
		return t.toString(16).toUpperCase().slice(1)
	}
	function yi(e) {
		var r = e[0] / 255,
		t = e[1] / 255,
		a = e[2] / 255;
		var n = Math.max(r, t, a),
		i = Math.min(r, t, a),
		s = n - i;
		if (s === 0) return [0, 0, r];
		var l = 0,
		o = 0,
		c = n + i;
		o = s / (c > 1 ? 2 - c: c);
		switch (n) {
		case r:
			l = ((t - a) / s + 6) % 6;
			break;
		case t:
			l = (a - r) / s + 2;
			break;
		case a:
			l = (r - t) / s + 4;
			break
		}
		return [l / 6, o, c / 2]
	}
	function xi(e) {
		var r = e[0],
		t = e[1],
		a = e[2];
		var n = t * 2 * (a < 0.5 ? a: 1 - a),
		i = a - n / 2;
		var s = [i, i, i],
		l = 6 * r;
		var o;
		if (t !== 0) switch (l | 0) {
		case 0:
		case 6:
			o = n * l;
			s[0] += n;
			s[1] += o;
			break;
		case 1:
			o = n * (2 - l);
			s[0] += o;
			s[1] += n;
			break;
		case 2:
			o = n * (l - 2);
			s[1] += n;
			s[2] += o;
			break;
		case 3:
			o = n * (4 - l);
			s[1] += o;
			s[2] += n;
			break;
		case 4:
			o = n * (l - 4);
			s[2] += n;
			s[0] += o;
			break;
		case 5:
			o = n * (6 - l);
			s[2] += o;
			s[0] += n;
			break
		}
		for (var c = 0; c != 3; ++c) s[c] = Math.round(s[c] * 255);
		return s
	}
	function Si(e, r) {
		if (r === 0) return e;
		var t = yi(wi(e));
		if (r < 0) t[2] = t[2] * (1 + r);
		else t[2] = 1 - (1 - t[2]) * (1 - r);
		return ki(xi(t))
	}
	var Ci = 6,
	_i = 15,
	Ai = 1,
	Ti = Ci;
	function Ei(e) {
		return Math.floor((e + Math.round(128 / Ti) / 256) * Ti)
	}
	function Fi(e) {
		return Math.floor(((e - 5) / Ti) * 100 + 0.5) / 100
	}
	function Di(e) {
		return Math.round(((e * Ti + 5) / Ti) * 256) / 256
	}
	function Oi(e) {
		return Di(Fi(Ei(e)))
	}
	function Mi(e) {
		var r = Math.abs(e - Oi(e)),
		t = Ti;
		if (r > 0.005) for (Ti = Ai; Ti < _i; ++Ti) if (Math.abs(e - Oi(e)) <= r) {
			r = Math.abs(e - Oi(e));
			t = Ti
		}
		Ti = t
	}
	function Ni(e) {
		if (e.width) {
			e.wpx = Ei(e.width);
			e.wch = Fi(e.wpx);
			e.MDW = Ti
		} else if (e.wpx) {
			e.wch = Fi(e.wpx);
			e.width = Di(e.wch);
			e.MDW = Ti
		} else if (typeof e.wch == "number") {
			e.width = Di(e.wch);
			e.wpx = Ei(e.width);
			e.MDW = Ti
		}
		if (e.customWidth) delete e.customWidth
	}
	var Ii = 96,
	Pi = Ii;
	function Ri(e) {
		return (e * 96) / Pi
	}
	function Li(e) {
		return (e * Pi) / 96
	}
	function Ui(e, r, t, a) {
		r.Borders = [];
		var n = {};
		var i = false; (e[0].match(qr) || []).forEach(function(e) {
			var t = rt(e);
			switch (tt(t[0])) {
			case "<borders":
			case "<borders>":
			case "</borders>":
				break;
			case "<border":
			case "<border>":
			case "<border/>":
				n = {};
				if (t.diagonalUp) n.diagonalUp = mt(t.diagonalUp);
				if (t.diagonalDown) n.diagonalDown = mt(t.diagonalDown);
				r.Borders.push(n);
				break;
			case "</border>":
				break;
			case "<left/>":
				break;
			case "<left":
			case "<left>":
				break;
			case "</left>":
				break;
			case "<right/>":
				break;
			case "<right":
			case "<right>":
				break;
			case "</right>":
				break;
			case "<top/>":
				break;
			case "<top":
			case "<top>":
				break;
			case "</top>":
				break;
			case "<bottom/>":
				break;
			case "<bottom":
			case "<bottom>":
				break;
			case "</bottom>":
				break;
			case "<diagonal":
			case "<diagonal>":
			case "<diagonal/>":
				break;
			case "</diagonal>":
				break;
			case "<horizontal":
			case "<horizontal>":
			case "<horizontal/>":
				break;
			case "</horizontal>":
				break;
			case "<vertical":
			case "<vertical>":
			case "<vertical/>":
				break;
			case "</vertical>":
				break;
			case "<start":
			case "<start>":
			case "<start/>":
				break;
			case "</start>":
				break;
			case "<end":
			case "<end>":
			case "<end/>":
				break;
			case "</end>":
				break;
			case "<color":
			case "<color>":
				break;
			case "<color/>":
			case "</color>":
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
				break;
			case "<ext":
				i = true;
				break;
			case "</ext>":
				i = false;
				break;
			default:
				if (a && a.WTF) {
					if (!i) throw new Error("unrecognized " + t[0] + " in borders");
				}
			}
		})
	}
	function zi(e, r, t, a) {
		r.Fills = [];
		var n = {};
		var i = false; (e[0].match(qr) || []).forEach(function(e) {
			var t = rt(e);
			switch (tt(t[0])) {
			case "<fills":
			case "<fills>":
			case "</fills>":
				break;
			case "<fill>":
			case "<fill":
			case "<fill/>":
				n = {};
				r.Fills.push(n);
				break;
			case "</fill>":
				break;
			case "<gradientFill>":
				break;
			case "<gradientFill":
			case "</gradientFill>":
				r.Fills.push(n);
				n = {};
				break;
			case "<patternFill":
			case "<patternFill>":
				if (t.patternType) n.patternType = t.patternType;
				break;
			case "<patternFill/>":
			case "</patternFill>":
				break;
			case "<bgColor":
				if (!n.bgColor) n.bgColor = {};
				if (t.indexed) n.bgColor.indexed = parseInt(t.indexed, 10);
				if (t.theme) n.bgColor.theme = parseInt(t.theme, 10);
				if (t.tint) n.bgColor.tint = parseFloat(t.tint);
				if (t.rgb) n.bgColor.rgb = t.rgb.slice(-6);
				break;
			case "<bgColor/>":
			case "</bgColor>":
				break;
			case "<fgColor":
				if (!n.fgColor) n.fgColor = {};
				if (t.theme) n.fgColor.theme = parseInt(t.theme, 10);
				if (t.tint) n.fgColor.tint = parseFloat(t.tint);
				if (t.rgb != null) n.fgColor.rgb = t.rgb.slice(-6);
				break;
			case "<fgColor/>":
			case "</fgColor>":
				break;
			case "<stop":
			case "<stop/>":
				break;
			case "</stop>":
				break;
			case "<color":
			case "<color/>":
				break;
			case "</color>":
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
				break;
			case "<ext":
				i = true;
				break;
			case "</ext>":
				i = false;
				break;
			default:
				if (a && a.WTF) {
					if (!i) throw new Error("unrecognized " + t[0] + " in fills");
				}
			}
		})
	}
	function $i(e, r, t, a) {
		r.Fonts = [];
		var n = {};
		var s = false; (e[0].match(qr) || []).forEach(function(e) {
			var l = rt(e);
			switch (tt(l[0])) {
			case "<fonts":
			case "<fonts>":
			case "</fonts>":
				break;
			case "<font":
			case "<font>":
				break;
			case "</font>":
			case "<font/>":
				r.Fonts.push(n);
				n = {};
				break;
			case "<name":
				if (l.val) n.name = kt(l.val);
				break;
			case "<name/>":
			case "</name>":
				break;
			case "<b":
				n.bold = l.val ? mt(l.val) : 1;
				break;
			case "<b/>":
				n.bold = 1;
				break;
			case "<i":
				n.italic = l.val ? mt(l.val) : 1;
				break;
			case "<i/>":
				n.italic = 1;
				break;
			case "<u":
				switch (l.val) {
				case "none":
					n.underline = 0;
					break;
				case "single":
					n.underline = 1;
					break;
				case "double":
					n.underline = 2;
					break;
				case "singleAccounting":
					n.underline = 33;
					break;
				case "doubleAccounting":
					n.underline = 34;
					break
				}
				break;
			case "<u/>":
				n.underline = 1;
				break;
			case "<strike":
				n.strike = l.val ? mt(l.val) : 1;
				break;
			case "<strike/>":
				n.strike = 1;
				break;
			case "<outline":
				n.outline = l.val ? mt(l.val) : 1;
				break;
			case "<outline/>":
				n.outline = 1;
				break;
			case "<shadow":
				n.shadow = l.val ? mt(l.val) : 1;
				break;
			case "<shadow/>":
				n.shadow = 1;
				break;
			case "<condense":
				n.condense = l.val ? mt(l.val) : 1;
				break;
			case "<condense/>":
				n.condense = 1;
				break;
			case "<extend":
				n.extend = l.val ? mt(l.val) : 1;
				break;
			case "<extend/>":
				n.extend = 1;
				break;
			case "<sz":
				if (l.val) n.sz = +l.val;
				break;
			case "<sz/>":
			case "</sz>":
				break;
			case "<vertAlign":
				if (l.val) n.vertAlign = l.val;
				break;
			case "<vertAlign/>":
			case "</vertAlign>":
				break;
			case "<family":
				if (l.val) n.family = parseInt(l.val, 10);
				break;
			case "<family/>":
			case "</family>":
				break;
			case "<scheme":
				if (l.val) n.scheme = l.val;
				break;
			case "<scheme/>":
			case "</scheme>":
				break;
			case "<charset":
				if (l.val == "1") break;
				l.codepage = i[parseInt(l.val, 10)];
				break;
			case "<color":
				if (!n.color) n.color = {};
				if (l.auto) n.color.auto = mt(l.auto);
				if (l.rgb) n.color.rgb = l.rgb.slice(-6);
				else if (l.indexed) {
					n.color.index = parseInt(l.indexed, 10);
					var o = wn[n.color.index];
					if (n.color.index == 81) o = wn[1];
					if (!o) o = wn[1];
					n.color.rgb = o[0].toString(16) + o[1].toString(16) + o[2].toString(16)
				} else if (l.theme) {
					n.color.theme = parseInt(l.theme, 10);
					if (l.tint) n.color.tint = parseFloat(l.tint);
					if (l.theme && t.themeElements && t.themeElements.clrScheme) {
						n.color.rgb = Si(t.themeElements.clrScheme[n.color.theme].rgb, n.color.tint || 0)
					}
				}
				break;
			case "<color/>":
			case "</color>":
				break;
			case "<AlternateContent":
				s = true;
				break;
			case "</AlternateContent>":
				s = false;
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
				break;
			case "<ext":
				s = true;
				break;
			case "</ext>":
				s = false;
				break;
			default:
				if (a && a.WTF) {
					if (!s) throw new Error("unrecognized " + l[0] + " in fonts");
				}
			}
		})
	}
	function Wi(e, r, t) {
		r.NumberFmt = [];
		var a = ir(q);
		for (var n = 0; n < a.length; ++n) r.NumberFmt[a[n]] = q[a[n]];
		var i = e[0].match(qr);
		if (!i) return;
		for (n = 0; n < i.length; ++n) {
			var s = rt(i[n]);
			switch (tt(s[0])) {
			case "<numFmts":
			case "</numFmts>":
			case "<numFmts/>":
			case "<numFmts>":
				break;
			case "<numFmt":
				{
					var l = it(kt(s.formatCode)),
					o = parseInt(s.numFmtId, 10);
					r.NumberFmt[o] = l;
					if (o > 0) {
						if (o > 392) {
							for (o = 392; o > 60; --o) if (r.NumberFmt[o] == null) break;
							r.NumberFmt[o] = l
						}
						Ke(l, o)
					}
				}
				break;
			case "</numFmt>":
				break;
			default:
				if (t.WTF) throw new Error("unrecognized " + s[0] + " in numFmts");
			}
		}
	}
	var Hi = ["numFmtId", "fillId", "fontId", "borderId", "xfId"];
	var Vi = ["applyAlignment", "applyBorder", "applyFill", "applyFont", "applyNumberFormat", "applyProtection", "pivotButton", "quotePrefix", ];
	function Xi(e, r, t) {
		r.CellXf = [];
		var a;
		var n = false; (e[0].match(qr) || []).forEach(function(e) {
			var i = rt(e),
			s = 0;
			switch (tt(i[0])) {
			case "<cellXfs":
			case "<cellXfs>":
			case "<cellXfs/>":
			case "</cellXfs>":
				break;
			case "<xf":
			case "<xf/>":
			case "<xf>":
				a = i;
				delete a[0];
				for (s = 0; s < Hi.length; ++s) if (a[Hi[s]]) a[Hi[s]] = parseInt(a[Hi[s]], 10);
				for (s = 0; s < Vi.length; ++s) if (a[Vi[s]]) a[Vi[s]] = mt(a[Vi[s]]);
				if (r.NumberFmt && a.numFmtId > 392) {
					for (s = 392; s > 60; --s) if (r.NumberFmt[a.numFmtId] == r.NumberFmt[s]) {
						a.numFmtId = s;
						break
					}
				}
				r.CellXf.push(a);
				break;
			case "</xf>":
				break;
			case "<alignment":
			case "<alignment/>":
			case "<alignment>":
				var l = {};
				if (i.vertical) l.vertical = i.vertical;
				if (i.horizontal) l.horizontal = i.horizontal;
				if (i.textRotation != null) l.textRotation = i.textRotation;
				if (i.indent) l.indent = i.indent;
				if (i.wrapText) l.wrapText = mt(i.wrapText);
				a.alignment = l;
				break;
			case "</alignment>":
				break;
			case "<protection":
			case "<protection>":
				break;
			case "</protection>":
			case "<protection/>":
				break;
			case "<AlternateContent":
			case "<AlternateContent>":
				n = true;
				break;
			case "</AlternateContent>":
				n = false;
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
				break;
			case "<ext":
				n = true;
				break;
			case "</ext>":
				n = false;
				break;
			default:
				if (t && t.WTF) {
					if (!n) throw new Error("unrecognized " + i[0] + " in cellXfs");
				}
			}
		})
	}
	var Yi = (function Jc() {
		var e = /<(?:\w+:)?numFmts([^>]*)>[\S\s]*?<\/(?:\w+:)?numFmts>/;
		var r = /<(?:\w+:)?cellXfs([^>]*)>[\S\s]*?<\/(?:\w+:)?cellXfs>/;
		var t = /<(?:\w+:)?fills([^>]*)>[\S\s]*?<\/(?:\w+:)?fills>/;
		var a = /<(?:\w+:)?fonts([^>]*)>[\S\s]*?<\/(?:\w+:)?fonts>/;
		var n = /<(?:\w+:)?borders([^>]*)>[\S\s]*?<\/(?:\w+:)?borders>/;
		return function i(s, l, o) {
			var c = {};
			if (!s) return c;
			s = s.replace(/<!--([\s\S]*?)-->/gm, "").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm, "");
			var f;
			if ((f = s.match(e))) Wi(f, c, o);
			if ((f = s.match(a))) $i(f, c, l, o);
			if ((f = s.match(t))) zi(f, c, l, o);
			if ((f = s.match(n))) Ui(f, c, l, o);
			if ((f = s.match(r))) Xi(f, c, o);
			return c
		}
	})();
	var Zi = ["</a:lt1>", "</a:dk1>", "</a:lt2>", "</a:dk2>", "</a:accent1>", "</a:accent2>", "</a:accent3>", "</a:accent4>", "</a:accent5>", "</a:accent6>", "</a:hlink>", "</a:folHlink>", ];
	function Ki(e, r, t) {
		r.themeElements.clrScheme = [];
		var a = {}; (e[0].match(qr) || []).forEach(function(e) {
			var n = rt(e);
			switch (n[0]) {
			case "<a:clrScheme":
			case "</a:clrScheme>":
				break;
			case "<a:srgbClr":
				a.rgb = n.val;
				break;
			case "<a:sysClr":
				a.rgb = n.lastClr;
				break;
			case "<a:dk1>":
			case "</a:dk1>":
			case "<a:lt1>":
			case "</a:lt1>":
			case "<a:dk2>":
			case "</a:dk2>":
			case "<a:lt2>":
			case "</a:lt2>":
			case "<a:accent1>":
			case "</a:accent1>":
			case "<a:accent2>":
			case "</a:accent2>":
			case "<a:accent3>":
			case "</a:accent3>":
			case "<a:accent4>":
			case "</a:accent4>":
			case "<a:accent5>":
			case "</a:accent5>":
			case "<a:accent6>":
			case "</a:accent6>":
			case "<a:hlink>":
			case "</a:hlink>":
			case "<a:folHlink>":
			case "</a:folHlink>":
				if (n[0].charAt(1) === "/") {
					r.themeElements.clrScheme[Zi.indexOf(n[0])] = a;
					a = {}
				} else {
					a.name = n[0].slice(3, n[0].length - 1)
				}
				break;
			default:
				if (t && t.WTF) throw new Error("Unrecognized " + n[0] + " in clrScheme");
			}
		})
	}
	function qi() {}
	function Qi() {}
	var es = /<a:clrScheme([^>]*)>[\s\S]*<\/a:clrScheme>/;
	var rs = /<a:fontScheme([^>]*)>[\s\S]*<\/a:fontScheme>/;
	var ts = /<a:fmtScheme([^>]*)>[\s\S]*<\/a:fmtScheme>/;
	function as(e, r, t) {
		r.themeElements = {};
		var a; [["clrScheme", es, Ki], ["fontScheme", rs, qi], ["fmtScheme", ts, Qi], ].forEach(function(n) {
			if (! (a = e.match(n[1]))) throw new Error(n[0] + " not found in themeElements");
			n[2](a, r, t)
		})
	}
	var ns = /<a:themeElements([^>]*)>[\s\S]*<\/a:themeElements>/;
	function is(e, r) {
		if (!e || e.length === 0) e = ss();
		var t;
		var a = {};
		if (! (t = e.match(ns))) throw new Error("themeElements not found in theme");
		as(t[0], a, r);
		a.raw = e;
		return a
	}
	function ss(e, r) {
		if (r && r.themeXLSX) return r.themeXLSX;
		if (e && typeof e.raw == "string") return e.raw;
		var t = [Yr];
		t[t.length] = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">';
		t[t.length] = "<a:themeElements>";
		t[t.length] = '<a:clrScheme name="Office">';
		t[t.length] = '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>';
		t[t.length] = '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>';
		t[t.length] = '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>';
		t[t.length] = '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>';
		t[t.length] = '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>';
		t[t.length] = '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>';
		t[t.length] = '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>';
		t[t.length] = '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>';
		t[t.length] = '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>';
		t[t.length] = '<a:accent6><a:srgbClr val="F79646"/></a:accent6>';
		t[t.length] = '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>';
		t[t.length] = '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>';
		t[t.length] = "</a:clrScheme>";
		t[t.length] = '<a:fontScheme name="Office">';
		t[t.length] = "<a:majorFont>";
		t[t.length] = '<a:latin typeface="Cambria"/>';
		t[t.length] = '<a:ea typeface=""/>';
		t[t.length] = '<a:cs typeface=""/>';
		t[t.length] = '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
		t[t.length] = '<a:font script="Hang" typeface="맑은 고딕"/>';
		t[t.length] = '<a:font script="Hans" typeface="宋体"/>';
		t[t.length] = '<a:font script="Hant" typeface="新細明體"/>';
		t[t.length] = '<a:font script="Arab" typeface="Times New Roman"/>';
		t[t.length] = '<a:font script="Hebr" typeface="Times New Roman"/>';
		t[t.length] = '<a:font script="Thai" typeface="Tahoma"/>';
		t[t.length] = '<a:font script="Ethi" typeface="Nyala"/>';
		t[t.length] = '<a:font script="Beng" typeface="Vrinda"/>';
		t[t.length] = '<a:font script="Gujr" typeface="Shruti"/>';
		t[t.length] = '<a:font script="Khmr" typeface="MoolBoran"/>';
		t[t.length] = '<a:font script="Knda" typeface="Tunga"/>';
		t[t.length] = '<a:font script="Guru" typeface="Raavi"/>';
		t[t.length] = '<a:font script="Cans" typeface="Euphemia"/>';
		t[t.length] = '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
		t[t.length] = '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
		t[t.length] = '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
		t[t.length] = '<a:font script="Thaa" typeface="MV Boli"/>';
		t[t.length] = '<a:font script="Deva" typeface="Mangal"/>';
		t[t.length] = '<a:font script="Telu" typeface="Gautami"/>';
		t[t.length] = '<a:font script="Taml" typeface="Latha"/>';
		t[t.length] = '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
		t[t.length] = '<a:font script="Orya" typeface="Kalinga"/>';
		t[t.length] = '<a:font script="Mlym" typeface="Kartika"/>';
		t[t.length] = '<a:font script="Laoo" typeface="DokChampa"/>';
		t[t.length] = '<a:font script="Sinh" typeface="Iskoola Pota"/>';
		t[t.length] = '<a:font script="Mong" typeface="Mongolian Baiti"/>';
		t[t.length] = '<a:font script="Viet" typeface="Times New Roman"/>';
		t[t.length] = '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
		t[t.length] = '<a:font script="Geor" typeface="Sylfaen"/>';
		t[t.length] = "</a:majorFont>";
		t[t.length] = "<a:minorFont>";
		t[t.length] = '<a:latin typeface="Calibri"/>';
		t[t.length] = '<a:ea typeface=""/>';
		t[t.length] = '<a:cs typeface=""/>';
		t[t.length] = '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
		t[t.length] = '<a:font script="Hang" typeface="맑은 고딕"/>';
		t[t.length] = '<a:font script="Hans" typeface="宋体"/>';
		t[t.length] = '<a:font script="Hant" typeface="新細明體"/>';
		t[t.length] = '<a:font script="Arab" typeface="Arial"/>';
		t[t.length] = '<a:font script="Hebr" typeface="Arial"/>';
		t[t.length] = '<a:font script="Thai" typeface="Tahoma"/>';
		t[t.length] = '<a:font script="Ethi" typeface="Nyala"/>';
		t[t.length] = '<a:font script="Beng" typeface="Vrinda"/>';
		t[t.length] = '<a:font script="Gujr" typeface="Shruti"/>';
		t[t.length] = '<a:font script="Khmr" typeface="DaunPenh"/>';
		t[t.length] = '<a:font script="Knda" typeface="Tunga"/>';
		t[t.length] = '<a:font script="Guru" typeface="Raavi"/>';
		t[t.length] = '<a:font script="Cans" typeface="Euphemia"/>';
		t[t.length] = '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
		t[t.length] = '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
		t[t.length] = '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
		t[t.length] = '<a:font script="Thaa" typeface="MV Boli"/>';
		t[t.length] = '<a:font script="Deva" typeface="Mangal"/>';
		t[t.length] = '<a:font script="Telu" typeface="Gautami"/>';
		t[t.length] = '<a:font script="Taml" typeface="Latha"/>';
		t[t.length] = '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
		t[t.length] = '<a:font script="Orya" typeface="Kalinga"/>';
		t[t.length] = '<a:font script="Mlym" typeface="Kartika"/>';
		t[t.length] = '<a:font script="Laoo" typeface="DokChampa"/>';
		t[t.length] = '<a:font script="Sinh" typeface="Iskoola Pota"/>';
		t[t.length] = '<a:font script="Mong" typeface="Mongolian Baiti"/>';
		t[t.length] = '<a:font script="Viet" typeface="Arial"/>';
		t[t.length] = '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
		t[t.length] = '<a:font script="Geor" typeface="Sylfaen"/>';
		t[t.length] = "</a:minorFont>";
		t[t.length] = "</a:fontScheme>";
		t[t.length] = '<a:fmtScheme name="Office">';
		t[t.length] = "<a:fillStyleLst>";
		t[t.length] = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
		t[t.length] = '<a:gradFill rotWithShape="1">';
		t[t.length] = "<a:gsLst>";
		t[t.length] = '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
		t[t.length] = '<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
		t[t.length] = '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		t[t.length] = "</a:gsLst>";
		t[t.length] = '<a:lin ang="16200000" scaled="1"/>';
		t[t.length] = "</a:gradFill>";
		t[t.length] = '<a:gradFill rotWithShape="1">';
		t[t.length] = "<a:gsLst>";
		t[t.length] = '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs>';
		t[t.length] = '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		t[t.length] = "</a:gsLst>";
		t[t.length] = '<a:lin ang="16200000" scaled="0"/>';
		t[t.length] = "</a:gradFill>";
		t[t.length] = "</a:fillStyleLst>";
		t[t.length] = "<a:lnStyleLst>";
		t[t.length] = '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>';
		t[t.length] = '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
		t[t.length] = '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
		t[t.length] = "</a:lnStyleLst>";
		t[t.length] = "<a:effectStyleLst>";
		t[t.length] = "<a:effectStyle>";
		t[t.length] = "<a:effectLst>";
		t[t.length] = '<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw>';
		t[t.length] = "</a:effectLst>";
		t[t.length] = "</a:effectStyle>";
		t[t.length] = "<a:effectStyle>";
		t[t.length] = "<a:effectLst>";
		t[t.length] = '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
		t[t.length] = "</a:effectLst>";
		t[t.length] = "</a:effectStyle>";
		t[t.length] = "<a:effectStyle>";
		t[t.length] = "<a:effectLst>";
		t[t.length] = '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
		t[t.length] = "</a:effectLst>";
		t[t.length] = '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d>';
		t[t.length] = '<a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>';
		t[t.length] = "</a:effectStyle>";
		t[t.length] = "</a:effectStyleLst>";
		t[t.length] = "<a:bgFillStyleLst>";
		t[t.length] = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
		t[t.length] = '<a:gradFill rotWithShape="1">';
		t[t.length] = "<a:gsLst>";
		t[t.length] = '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		t[t.length] = '<a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		t[t.length] = '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>';
		t[t.length] = "</a:gsLst>";
		t[t.length] = '<a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>';
		t[t.length] = "</a:gradFill>";
		t[t.length] = '<a:gradFill rotWithShape="1">';
		t[t.length] = "<a:gsLst>";
		t[t.length] = '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
		t[t.length] = '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>';
		t[t.length] = "</a:gsLst>";
		t[t.length] = '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>';
		t[t.length] = "</a:gradFill>";
		t[t.length] = "</a:bgFillStyleLst>";
		t[t.length] = "</a:fmtScheme>";
		t[t.length] = "</a:themeElements>";
		t[t.length] = "<a:objectDefaults>";
		t[t.length] = "<a:spDef>";
		t[t.length] = '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style>';
		t[t.length] = "</a:spDef>";
		t[t.length] = "<a:lnDef>";
		t[t.length] = '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style>';
		t[t.length] = "</a:lnDef>";
		t[t.length] = "</a:objectDefaults>";
		t[t.length] = "<a:extraClrSchemeLst/>";
		t[t.length] = "</a:theme>";
		return t.join("")
	}
	function ls(e, r, t) {
		var a = {
			Types: [],
			Cell: [],
			Value: [],
		};
		if (!e) return a;
		var n = false;
		var i = 2;
		var s;
		e.replace(qr,
		function(e) {
			var r = rt(e);
			switch (tt(r[0])) {
			case "<?xml":
				break;
			case "<metadata":
			case "</metadata>":
				break;
			case "<metadataTypes":
			case "</metadataTypes>":
				break;
			case "<metadataType":
				a.Types.push({
					name:
					r.name,
				});
				break;
			case "</metadataType>":
				break;
			case "<futureMetadata":
				for (var l = 0; l < a.Types.length; ++l) if (a.Types[l].name == r.name) s = a.Types[l];
				break;
			case "</futureMetadata>":
				break;
			case "<bk>":
				break;
			case "</bk>":
				break;
			case "<rc":
				if (i == 1) a.Cell.push({
					type: a.Types[r.t - 1].name,
					index: +r.v,
				});
				else if (i == 0) a.Value.push({
					type: a.Types[r.t - 1].name,
					index: +r.v,
				});
				break;
			case "</rc>":
				break;
			case "<cellMetadata":
				i = 1;
				break;
			case "</cellMetadata>":
				i = 2;
				break;
			case "<valueMetadata":
				i = 0;
				break;
			case "</valueMetadata>":
				i = 2;
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
			case "<extLst/>":
				break;
			case "<ext":
				n = true;
				break;
			case "</ext>":
				n = false;
				break;
			case "<rvb":
				if (!s) break;
				if (!s.offsets) s.offsets = [];
				s.offsets.push( + r.i);
				break;
			default:
				if (!n && (t == null ? void 0 : t.WTF)) throw new Error("unrecognized " + r[0] + " in metadata");
			}
			return e
		});
		return a
	}
	function cs() {}
	function fs(e, r, t, a) {
		if (!e) return e;
		var n = a || {};
		var i = false,
		s = false;
		Ca(e,
		function l(e, r, t) {
			if (s) return;
			switch (t) {
			case 359:
			case 363:
			case 364:
			case 366:
			case 367:
			case 368:
			case 369:
			case 370:
			case 371:
			case 472:
			case 577:
			case 578:
			case 579:
			case 580:
			case 581:
			case 582:
			case 583:
			case 584:
			case 585:
			case 586:
			case 587:
				break;
			case 35:
				i = true;
				break;
			case 36:
				i = false;
				break;
			default:
				if (r.T) {} else if (!i || n.WTF) throw new Error("Unexpected record 0x" + t.toString(16));
			}
		},
		n)
	}
	function us(e, r) {
		if (!e) return "??";
		var t = (e.match(/<c:chart [^>]*r:id="([^"]*)"/) || ["", ""])[1];
		return r["!id"][t].Target
	}
	var hs = /<(?:\w+:)?shape(?:[^\w][^>]*)?>([\s\S]*?)<\/(?:\w+:)?shape>/g;
	function ds(e, r, t) {
		var a = 0; (e.match(hs) || []).forEach(function(e) {
			var n = "";
			var i = true;
			var s = -1;
			var l = -1,
			o = -1;
			e.replace(qr,
			function(r, t) {
				var a = rt(r);
				switch (tt(a[0])) {
				case "<ClientData":
					if (a.ObjectType) n = a.ObjectType;
					break;
				case "<Visible":
				case "<Visible/>":
					i = false;
					break;
				case "<Row":
				case "<Row>":
					s = t + r.length;
					break;
				case "</Row>":
					l = +e.slice(s, t).trim();
					break;
				case "<Column":
				case "<Column>":
					s = t + r.length;
					break;
				case "</Column>":
					o = +e.slice(s, t).trim();
					break
				}
				return ""
			});
			switch (n) {
			case "Note":
				var c = yc(r, l >= 0 && o >= 0 ? $a({
					r: l,
					c: o,
				}) : t[a].ref);
				if (c.c) {
					c.c.hidden = i
				}++a;
				break
			}
		})
	}
	function vs(e, r, t, a) {
		var n = e["!data"] != null;
		var i;
		r.forEach(function(r) {
			var s = za(r.ref);
			if (s.r < 0 || s.c < 0) return;
			if (n) {
				if (!e["!data"][s.r]) e["!data"][s.r] = [];
				i = e["!data"][s.r][s.c]
			} else i = e[r.ref];
			if (!i) {
				i = {
					t: "z",
				};
				if (n) e["!data"][s.r][s.c] = i;
				else e[r.ref] = i;
				var l = Xa(e["!ref"] || "BDWGO1000001:A1");
				if (l.s.r > s.r) l.s.r = s.r;
				if (l.e.r < s.r) l.e.r = s.r;
				if (l.s.c > s.c) l.s.c = s.c;
				if (l.e.c < s.c) l.e.c = s.c;
				var o = ja(l);
				e["!ref"] = o
			}
			if (!i.c) i.c = [];
			var c = {
				a: r.author,
				t: r.t,
				r: r.r,
				T: t,
			};
			if (r.h) c.h = r.h;
			for (var f = i.c.length - 1; f >= 0; --f) {
				if (!t && i.c[f].T) return;
				if (t && !i.c[f].T) i.c.splice(f, 1)
			}
			if (t && a) for (f = 0; f < a.length; ++f) {
				if (c.a == a[f].id) {
					c.a = a[f].name || c.a;
					break
				}
			}
			i.c.push(c)
		})
	}
	function gs(e, r) {
		if (e.match(/<(?:\w+:)?comments *\/>/)) return [];
		var t = [];
		var a = [];
		var n = e.match(/<(?:\w+:)?authors>([\s\S]*)<\/(?:\w+:)?authors>/);
		if (n && n[1]) n[1].split(/<\/\w*:?author>/).forEach(function(e) {
			if (e === "" || e.trim() === "") return;
			var r = e.match(/<(?:\w+:)?author[^>]*>(.*)/);
			if (r) t.push(r[1])
		});
		var i = e.match(/<(?:\w+:)?commentList>([\s\S]*)<\/(?:\w+:)?commentList>/);
		if (i && i[1]) i[1].split(/<\/\w*:?comment>/).forEach(function(e) {
			if (e === "" || e.trim() === "") return;
			var n = e.match(/<(?:\w+:)?comment[^>]*>/);
			if (!n) return;
			var i = rt(n[0]);
			var s = {
				author: (i.authorId && t[i.authorId]) || "sheetjsghost",
				ref: i.ref,
				guid: i.guid,
			};
			var l = za(i.ref);
			if (r.sheetRows && r.sheetRows <= l.r) return;
			var o = e.match(/<(?:\w+:)?text>([\s\S]*)<\/(?:\w+:)?text>/);
			var c = ( !! o && !!o[1] && hi(o[1])) || {
				r: "",
				t: "",
				h: "",
			};
			s.r = c.r;
			if (c.r == "<t></t>") c.t = c.h = "";
			s.t = (c.t || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
			if (r.cellHTML) s.h = c.h;
			a.push(s)
		});
		return a
	}
	function ws(e, r) {
		var t = [];
		var a = false,
		n = {},
		i = 0;
		e.replace(qr,
		function s(l, o) {
			var c = rt(l);
			switch (tt(c[0])) {
			case "<?xml":
				break;
			case "<ThreadedComments":
				break;
			case "</ThreadedComments>":
				break;
			case "<threadedComment":
				n = {
					author: c.personId,
					guid: c.id,
					ref: c.ref,
					T: 1,
				};
				break;
			case "</threadedComment>":
				if (n.t != null) t.push(n);
				break;
			case "<text>":
			case "<text":
				i = o + l.length;
				break;
			case "</text>":
				n.t = e.slice(i, o).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
				break;
			case "<mentions":
			case "<mentions>":
				a = true;
				break;
			case "</mentions>":
				a = false;
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
			case "<extLst/>":
				break;
			case "<ext":
				a = true;
				break;
			case "</ext>":
				a = false;
				break;
			default:
				if (!a && r.WTF) throw new Error("unrecognized " + c[0] + " in threaded comments");
			}
			return l
		});
		return t
	}
	function ys(e, r) {
		var t = [];
		var a = false;
		e.replace(qr,
		function n(e) {
			var n = rt(e);
			switch (tt(n[0])) {
			case "<?xml":
				break;
			case "<personList":
				break;
			case "</personList>":
				break;
			case "<person":
				t.push({
					name:
					n.displayname,
					id: n.id,
				});
				break;
			case "</person>":
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
			case "<extLst/>":
				break;
			case "<ext":
				a = true;
				break;
			case "</ext>":
				a = false;
				break;
			default:
				if (!a && r.WTF) throw new Error("unrecognized " + n[0] + " in threaded comments");
			}
			return e
		});
		return t
	}
	var Ss = "application/vnd.ms-office.vbaProject";
	function _s(e, r) {
		r.FullPaths.forEach(function(t, a) {
			if (a == 0) return;
			var n = t.replace(/[^\/]*[\/]/, "/_VBA_PROJECT_CUR/");
			if (n.slice(-1) !== "/") Qe.utils.cfb_add(e, n, r.FileIndex[a].content)
		})
	}
	function Ts() {
		return {
			"!type": "dialog",
		}
	}
	function Es() {
		return {
			"!type": "dialog",
		}
	}
	function Fs() {
		return {
			"!type": "macro",
		}
	}
	function Ds() {
		return {
			"!type": "macro",
		}
	}
	var Os = (function() {
		var e = /(^|[^A-Za-z_])R(\[?-?\d+\]|[1-9]\d*|)C(\[?-?\d+\]|[1-9]\d*|)(?![A-Za-z0-9_])/g;
		var r = {
			r: 0,
			c: 0,
		};
		function t(e, t, a, n) {
			var i = false,
			s = false;
			if (a.length == 0) s = true;
			else if (a.charAt(0) == "[") {
				s = true;
				a = a.slice(1, -1)
			}
			if (n.length == 0) i = true;
			else if (n.charAt(0) == "[") {
				i = true;
				n = n.slice(1, -1)
			}
			var l = a.length > 0 ? parseInt(a, 10) | 0 : 0,
			o = n.length > 0 ? parseInt(n, 10) | 0 : 0;
			if (i) o += r.c;
			else--o;
			if (s) l += r.r;
			else--l;
			return t + (i ? "": "$") + Ra(o) + (s ? "": "$") + Ma(l)
		}
		return function a(n, i) {
			r = i;
			return n.replace(e, t)
		}
	})();
	var Ms = /(^|[^._A-Z0-9])(\$?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])(\$?)(\d{1,7})(?![_.\(A-Za-z0-9])/g;
	try {
		Ms = /(^|[^._A-Z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)(10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})(?![_.\(A-Za-z0-9])/g
	} catch(Ns) {}
	var Is = (function() {
		return function e(r, t) {
			return r.replace(Ms,
			function(e, r, a, n, i, s) {
				var l = Pa(n) - (a ? 0 : t.c);
				var o = Oa(s) - (i ? 0 : t.r);
				var c = i == "$" ? o + 1 : o == 0 ? "": "[" + o + "]";
				var f = a == "$" ? l + 1 : l == 0 ? "": "[" + l + "]";
				return r + "R" + c + "C" + f
			})
		}
	})();
	function Ps(e, r) {
		return e.replace(Ms,
		function(e, t, a, n, i, s) {
			return (t + (a == "$" ? a + n: Ra(Pa(n) + r.c)) + (i == "$" ? i + s: Ma(Oa(s) + r.r)))
		})
	}
	function Rs(e, r, t) {
		var a = Wa(r),
		n = a.s,
		i = za(t);
		var s = {
			r: i.r - n.r,
			c: i.c - n.c,
		};
		return Ps(e, s)
	}
	function Ls(e) {
		if (e.length == 1) return false;
		return true
	}
	function Bs(e) {
		return e.replace(/_xlfn\./g, "")
	}
	function Us(e) {
		if (e.slice(0, 3) == "of:") e = e.slice(3);
		if (e.charCodeAt(0) == 61) {
			e = e.slice(1);
			if (e.charCodeAt(0) == 61) e = e.slice(1)
		}
		e = e.replace(/COM\.MICROSOFT\./g, "");
		e = e.replace(/\[((?:\.[A-Z]+[0-9]+)(?::\.[A-Z]+[0-9]+)?)\]/g,
		function(e, r) {
			return r.replace(/\./g, "")
		});
		e = e.replace(/\$'([^']|'')+'/g,
		function(e) {
			return e.slice(1)
		});
		e = e.replace(/\$([^\]\. #$]+)/g,
		function(e, r) {
			return r.match(/^([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])?(10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})?$/) ? e: r
		});
		e = e.replace(/\[.(#[A-Z]*[?!])\]/g, "$1");
		return e.replace(/[;~]/g, ",").replace(/\|/g, ";")
	}
	function $s(e) {
		e = e.replace(/\$'([^']|'')+'/g,
		function(e) {
			return e.slice(1)
		});
		e = e.replace(/\$([^\]\. #$]+)/g,
		function(e, r) {
			return r.match(/^([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])?(10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})?$/) ? e: r
		});
		var r = e.split(":");
		var t = r[0].split(".")[0];
		return [t, r[0].split(".")[1] + (r.length > 1 ? ":" + (r[1].split(".")[1] || r[1].split(".")[0]) : ""), ]
	}
	var js = {};
	var Hs = {};
	function Zs(e, r, t, a, n, i, s) {
		try {
			if (a.cellNF) e.z = q[r]
		} catch(l) {
			if (a.WTF) throw l;
		}
		if (e.t === "z" && !a.cellStyles) return;
		if (e.t === "d" && typeof e.v === "string") e.v = wr(e.v);
		if ((!a || a.cellText !== false) && e.t !== "z") try {
			if (q[r] == null) Ke(Xe[r] || "General", r);
			if (e.t === "e") e.w = e.w || kn[e.v];
			else if (r === 0) {
				if (e.t === "n") {
					if ((e.v | 0) === e.v) e.w = e.v.toString(10);
					else e.w = oe(e.v)
				} else if (e.t === "d") {
					var o = dr(e.v, !!s);
					if ((o | 0) === o) e.w = o.toString(10);
					else e.w = oe(o)
				} else if (e.v === undefined) return "";
				else e.w = ce(e.v, Hs)
			} else if (e.t === "d") e.w = $e(r, dr(e.v, !!s), Hs);
			else e.w = $e(r, e.v, Hs)
		} catch(l) {
			if (a.WTF) throw l;
		}
		if (!a.cellStyles) return;
		if (t != null) try {
			e.s = i.Fills[t];
			if (e.s.fgColor && e.s.fgColor.theme && !e.s.fgColor.rgb) {
				e.s.fgColor.rgb = Si(n.themeElements.clrScheme[e.s.fgColor.theme].rgb, e.s.fgColor.tint || 0);
				if (a.WTF) e.s.fgColor.raw_rgb = n.themeElements.clrScheme[e.s.fgColor.theme].rgb
			}
			if (e.s.bgColor && e.s.bgColor.theme) {
				e.s.bgColor.rgb = Si(n.themeElements.clrScheme[e.s.bgColor.theme].rgb, e.s.bgColor.tint || 0);
				if (a.WTF) e.s.bgColor.raw_rgb = n.themeElements.clrScheme[e.s.bgColor.theme].rgb
			}
		} catch(l) {
			if (a.WTF && i.Fills) throw l;
		}
	}
	function qs(e, r) {
		var t = Xa(r);
		if (t.s.r <= t.e.r && t.s.c <= t.e.c && t.s.r >= 0 && t.s.c >= 0) e["!ref"] = ja(t)
	}
	var Qs = /<(?:\w:)?mergeCell ref=["'][A-Z0-9:]+['"]\s*[\/]?>/g;
	var el = /<(?:\w+:)?sheetData[^>]*>([\s\S]*)<\/(?:\w+:)?sheetData>/;
	var rl = /<(?:\w:)?hyperlink [^>]*>/gm;
	var tl = /"(\w*:\w*)"/;
	var al = /<(?:\w:)?col\b[^>]*[\/]?>/g;
	var nl = /<(?:\w:)?autoFilter[^>]*([\/]|>([\s\S]*)<\/(?:\w:)?autoFilter)>/g;
	var il = /<(?:\w:)?pageMargins[^>]*\/>/g;
	var sl = /<(?:\w:)?sheetPr\b(?:[^>a-z][^>]*)?\/>/;
	var ll = /<(?:\w:)?sheetPr[^>]*(?:[\/]|>([\s\S]*)<\/(?:\w:)?sheetPr)>/;
	var ol = /<(?:\w:)?sheetViews[^>]*(?:[\/]|>([\s\S]*)<\/(?:\w:)?sheetViews)>/;
	function cl(e, r, t, a, n, i, s) {
		if (!e) return e;
		if (!a) a = {
			"!id": {},
		};
		if (b != null && r.dense == null) r.dense = b;
		var l = {};
		if (r.dense) l["!data"] = [];
		var o = {
			s: {
				r: 2e6,
				c: 2e6,
			},
			e: {
				r: 0,
				c: 0,
			},
		};
		var c = "",
		f = "";
		var u = e.match(el);
		if (u) {
			c = e.slice(0, u.index);
			f = e.slice(u.index + u[0].length)
		} else c = f = e;
		var h = c.match(sl);
		if (h) ul(h[0], l, n, t);
		else if ((h = c.match(ll))) hl(h[0], h[1] || "", l, n, t, s, i);
		var d = (c.match(/<(?:\w*:)?dimension/) || {
			index: -1,
		}).index;
		if (d > 0) {
			var p = c.slice(d, d + 50).match(tl);
			if (p && !(r && r.nodim)) qs(l, p[1])
		}
		var m = c.match(ol);
		if (m && m[1]) _l(m[1], n);
		var v = [];
		if (r.cellStyles) {
			var g = c.match(al);
			if (g) kl(v, g)
		}
		if (u) El(u[1], l, r, o, i, s, n);
		var w = f.match(nl);
		if (w) l["!autofilter"] = xl(w[0]);
		var k = [];
		var y = f.match(Qs);
		if (y) for (d = 0; d != y.length; ++d) k[d] = Xa(y[d].slice(y[d].indexOf('"') + 1));
		var x = f.match(rl);
		if (x) gl(l, x, a);
		var S = f.match(il);
		if (S) l["!margins"] = bl(rt(S[0]));
		var C;
		if ((C = f.match(/legacyDrawing r:id="(.*?)"/))) l["!legrel"] = C[1];
		if (r && r.nodim) o.s.c = o.s.r = 0;
		if (!l["!ref"] && o.e.c >= o.s.c && o.e.r >= o.s.r) l["!ref"] = ja(o);
		if (r.sheetRows > 0 && l["!ref"]) {
			var _ = Xa(l["!ref"]);
			if (r.sheetRows <= +_.e.r) {
				_.e.r = r.sheetRows - 1;
				if (_.e.r > o.e.r) _.e.r = o.e.r;
				if (_.e.r < _.s.r) _.s.r = _.e.r;
				if (_.e.c > o.e.c) _.e.c = o.e.c;
				if (_.e.c < _.s.c) _.s.c = _.e.c;
				l["!fullref"] = l["!ref"];
				l["!ref"] = ja(_)
			}
		}
		if (v.length > 0) l["!cols"] = v;
		if (k.length > 0) l["!merges"] = k;
		if (a["!id"][l["!legrel"]]) l["!legdrawel"] = a["!id"][l["!legrel"]];
		return l
	}
	function ul(e, r, t, a) {
		var n = rt(e);
		if (!t.Sheets[a]) t.Sheets[a] = {};
		if (n.codeName) t.Sheets[a].CodeName = it(kt(n.codeName))
	}
	function hl(e, r, t, a, n) {
		ul(e.slice(0, e.indexOf(">")), t, a, n)
	}
	function gl(e, r, t) {
		var a = e["!data"] != null;
		for (var n = 0; n != r.length; ++n) {
			var i = rt(kt(r[n]), true);
			if (!i.ref) return;
			var s = ((t || {})["!id"] || [])[i.id];
			if (s) {
				i.Target = s.Target;
				if (i.location) i.Target += "#" + it(i.location)
			} else {
				i.Target = "#" + it(i.location);
				s = {
					Target: i.Target,
					TargetMode: "Internal",
				}
			}
			i.Rel = s;
			if (i.tooltip) {
				i.Tooltip = i.tooltip;
				delete i.tooltip
			}
			var l = Xa(i.ref);
			for (var o = l.s.r; o <= l.e.r; ++o) for (var c = l.s.c; c <= l.e.c; ++c) {
				var f = Ra(c) + Ma(o);
				if (a) {
					if (!e["!data"][o]) e["!data"][o] = [];
					if (!e["!data"][o][c]) e["!data"][o][c] = {
						t: "z",
						v: undefined,
					};
					e["!data"][o][c].l = i
				} else {
					if (!e[f]) e[f] = {
						t: "z",
						v: undefined,
					};
					e[f].l = i
				}
			}
		}
	}
	function bl(e) {
		var r = {}; ["left", "right", "top", "bottom", "header", "footer"].forEach(function(t) {
			if (e[t]) r[t] = parseFloat(e[t])
		});
		return r
	}
	function kl(e, r) {
		var t = false;
		for (var a = 0; a != r.length; ++a) {
			var n = rt(r[a], true);
			if (n.hidden) n.hidden = mt(n.hidden);
			var i = parseInt(n.min, 10) - 1,
			s = parseInt(n.max, 10) - 1;
			if (n.outlineLevel) n.level = +n.outlineLevel || 0;
			delete n.min;
			delete n.max;
			n.width = +n.width;
			if (!t && n.width) {
				t = true;
				Mi(n.width)
			}
			Ni(n);
			while (i <= s) e[i++] = yr(n)
		}
	}
	function xl(e) {
		var r = {
			ref: (e.match(/ref="([^"]*)"/) || [])[1],
		};
		return r
	}
	var Cl = /<(?:\w:)?sheetView(?:[^>a-z][^>]*)?\/?>/g;
	function _l(e, r) {
		if (!r.Views) r.Views = [{}]; (e.match(Cl) || []).forEach(function(e, t) {
			var a = rt(e);
			if (!r.Views[t]) r.Views[t] = {};
			if ( + a.zoomScale) r.Views[t].zoom = +a.zoomScale;
			if (a.rightToLeft && mt(a.rightToLeft)) r.Views[t].RTL = true
		})
	}
	var El = (function() {
		var e = /<(?:\w+:)?c[ \/>]/,
		r = /<\/(?:\w+:)?row>/;
		var t = /r=["']([^"']*)["']/,
		a = /<(?:\w+:)?is>([\S\s]*?)<\/(?:\w+:)?is>/;
		var n = /ref=["']([^"']*)["']/;
		var i = xt("v"),
		s = xt("f");
		return function l(o, c, f, u, h, d, p) {
			var m = 0,
			v = "",
			g = [],
			b = [],
			w = 0,
			k = 0,
			y = 0,
			x = "",
			S;
			var C, _ = 0,
			A = 0;
			var T, E;
			var F = 0,
			D = 0;
			var O = Array.isArray(d.CellXf),
			M;
			var N = [];
			var I = [];
			var P = c["!data"] != null;
			var R = [],
			L = {},
			B = false;
			var U = !!f.sheetStubs;
			var z = !!((p || {}).WBProps || {}).date1904;
			for (var $ = o.split(r), W = 0, j = $.length; W != j; ++W) {
				v = $[W].trim();
				var H = v.length;
				if (H === 0) continue;
				var V = 0;
				e: for (m = 0; m < H; ++m) switch (v[m]) {
				case ">":
					if (v[m - 1] != "/") {++m;
						break e
					}
					if (f && f.cellStyles) {
						C = rt(v.slice(V, m), true);
						_ = C.r != null ? parseInt(C.r, 10) : _ + 1;
						A = -1;
						if (f.sheetRows && f.sheetRows < _) continue;
						L = {};
						B = false;
						if (C.ht) {
							B = true;
							L.hpt = parseFloat(C.ht);
							L.hpx = Li(L.hpt)
						}
						if (C.hidden && mt(C.hidden)) {
							B = true;
							L.hidden = true
						}
						if (C.outlineLevel != null) {
							B = true;
							L.level = +C.outlineLevel
						}
						if (B) R[_ - 1] = L
					}
					break;
				case "<":
					V = m;
					break
				}
				if (V >= m) break;
				C = rt(v.slice(V, m), true);
				_ = C.r != null ? parseInt(C.r, 10) : _ + 1;
				A = -1;
				if (f.sheetRows && f.sheetRows < _) continue;
				if (!f.nodim) {
					if (u.s.r > _ - 1) u.s.r = _ - 1;
					if (u.e.r < _ - 1) u.e.r = _ - 1
				}
				if (f && f.cellStyles) {
					L = {};
					B = false;
					if (C.ht) {
						B = true;
						L.hpt = parseFloat(C.ht);
						L.hpx = Li(L.hpt)
					}
					if (C.hidden && mt(C.hidden)) {
						B = true;
						L.hidden = true
					}
					if (C.outlineLevel != null) {
						B = true;
						L.level = +C.outlineLevel
					}
					if (B) R[_ - 1] = L
				}
				g = v.slice(m).split(e);
				for (var X = 0; X != g.length; ++X) if (g[X].trim().charAt(0) != "<") break;
				g = g.slice(X);
				for (m = 0; m != g.length; ++m) {
					v = g[m].trim();
					if (v.length === 0) continue;
					b = v.match(t);
					w = m;
					k = 0;
					y = 0;
					v = "<c " + (v.slice(0, 1) == "<" ? ">": "") + v;
					if (b != null && b.length === 2) {
						w = 0;
						x = b[1];
						for (k = 0; k != x.length; ++k) {
							if ((y = x.charCodeAt(k) - 64) < 1 || y > 26) break;
							w = 26 * w + y
						}--w;
						A = w
					} else++A;
					for (k = 0; k != v.length; ++k) if (v.charCodeAt(k) === 62) break; ++k;
					C = rt(v.slice(0, k), true);
					if (!C.r) C.r = $a({
						r: _ - 1,
						c: A,
					});
					x = v.slice(k);
					S = {
						t: "",
					};
					if ((b = x.match(i)) != null && b[1] !== "") S.v = it(b[1]);
					if (f.cellFormula) {
						if ((b = x.match(s)) != null) {
							if (b[1] == "") {
								if (b[0].indexOf('t="shared"') > -1) {
									E = rt(b[0]);
									if (I[E.si]) S.f = Rs(I[E.si][1], I[E.si][2], C.r)
								}
							} else {
								S.f = it(kt(b[1]), true);
								if (!f.xlfn) S.f = Bs(S.f);
								if (b[0].indexOf('t="array"') > -1) {
									S.F = (x.match(n) || [])[1];
									if (S.F.indexOf(":") > -1) N.push([Xa(S.F), S.F])
								} else if (b[0].indexOf('t="shared"') > -1) {
									E = rt(b[0]);
									var G = it(kt(b[1]));
									if (!f.xlfn) G = Bs(G);
									I[parseInt(E.si, 10)] = [E, G, C.r]
								}
							}
						} else if ((b = x.match(/<f[^>]*\/>/))) {
							E = rt(b[0]);
							if (I[E.si]) S.f = Rs(I[E.si][1], I[E.si][2], C.r)
						}
						var Y = za(C.r);
						for (k = 0; k < N.length; ++k) if (Y.r >= N[k][0].s.r && Y.r <= N[k][0].e.r) if (Y.c >= N[k][0].s.c && Y.c <= N[k][0].e.c) S.F = N[k][1]
					}
					if (C.t == null && S.v === undefined) {
						if (S.f || S.F) {
							S.v = 0;
							S.t = "n"
						} else if (!U) continue;
						else S.t = "z"
					} else S.t = C.t || "n";
					if (u.s.c > A) u.s.c = A;
					if (u.e.c < A) u.e.c = A;
					switch (S.t) {
					case "n":
						if (S.v == "" || S.v == null) {
							if (!U) continue;
							S.t = "z"
						} else S.v = parseFloat(S.v);
						break;
					case "s":
						if (typeof S.v == "undefined") {
							if (!U) continue;
							S.t = "z"
						} else {
							T = js[parseInt(S.v, 10)];
							S.v = T.t;
							S.r = T.r;
							if (f.cellHTML) S.h = T.h
						}
						break;
					case "str":
						S.t = "s";
						S.v = S.v != null ? it(kt(S.v), true) : "";
						if (f.cellHTML) S.h = ut(S.v);
						break;
					case "inlineStr":
						b = x.match(a);
						S.t = "s";
						if (b != null && (T = hi(b[1]))) {
							S.v = T.t;
							if (f.cellHTML) S.h = T.h
						} else S.v = "";
						break;
					case "b":
						S.v = mt(S.v);
						break;
					case "d":
						if (f.cellDates) S.v = wr(S.v, z);
						else {
							S.v = dr(wr(S.v, z), z);
							S.t = "n"
						}
						break;
					case "e":
						if (!f || f.cellText !== false) S.w = S.v;
						S.v = yn[S.v];
						break
					}
					F = D = 0;
					M = null;
					if (O && C.s !== undefined) {
						M = d.CellXf[C.s];
						if (M != null) {
							if (M.numFmtId != null) F = M.numFmtId;
							if (f.cellStyles) {
								if (M.fillId != null) D = M.fillId
							}
						}
					}
					Zs(S, F, D, f, h, d, z);
					if (f.cellDates && O && S.t == "n" && Re(q[F])) {
						S.v = pr(S.v + (z ? 1462 : 0));
						S.t = typeof S.v == "number" ? "n": "d"
					}
					if (C.cm && f.xlmeta) {
						var J = (f.xlmeta.Cell || [])[ + C.cm - 1];
						if (J && J.type == "XLDAPR") S.D = true
					}
					var Z;
					if (f.nodim) {
						Z = za(C.r);
						if (u.s.r > Z.r) u.s.r = Z.r;
						if (u.e.r < Z.r) u.e.r = Z.r
					}
					if (P) {
						Z = za(C.r);
						if (!c["!data"][Z.r]) c["!data"][Z.r] = [];
						c["!data"][Z.r][Z.c] = S
					} else c[C.r] = S
				}
			}
			if (R.length > 0) c["!rows"] = R
		}
	})();
	function Ol(e) {
		var r = [];
		var t = e.match(/^<c:numCache>/);
		var a; (e.match(/<c:pt idx="(\d*)">(.*?)<\/c:pt>/gm) || []).forEach(function(e) {
			var a = e.match(/<c:pt idx="(\d*?)"><c:v>(.*)<\/c:v><\/c:pt>/);
			if (!a) return;
			r[ + a[1]] = t ? +a[2] : a[2]
		});
		var n = it((e.match(/<c:formatCode>([\s\S]*?)<\/c:formatCode>/) || ["", "General", ])[1]); (e.match(/<c:f>(.*?)<\/c:f>/gm) || []).forEach(function(e) {
			a = e.replace(/<.*?>/g, "")
		});
		return [r, n, a]
	}
	function Ml(e, r, t, a, n, i) {
		var s = i || {
			"!type": "chart",
		};
		if (!e) return i;
		var l = 0,
		o = 0,
		c = "A";
		var f = {
			s: {
				r: 2e6,
				c: 2e6,
			},
			e: {
				r: 0,
				c: 0,
			},
		}; (e.match(/<c:numCache>[\s\S]*?<\/c:numCache>/gm) || []).forEach(function(e) {
			var r = Ol(e);
			f.s.r = f.s.c = 0;
			f.e.c = l;
			c = Ra(l);
			r[0].forEach(function(e, t) {
				if (s["!data"]) {
					if (!s["!data"][t]) s["!data"][t] = [];
					s["!data"][t][l] = {
						t: "n",
						v: e,
						z: r[1],
					}
				} else s[c + Ma(t)] = {
					t: "n",
					v: e,
					z: r[1],
				};
				o = t
			});
			if (f.e.r < o) f.e.r = o; ++l
		});
		if (l > 0) s["!ref"] = ja(f);
		return s
	}
	function Nl(e, r, t, a, n) {
		if (!e) return e;
		if (!a) a = {
			"!id": {},
		};
		var i = {
			"!type": "chart",
			"!drawel": null,
			"!rel": "",
		};
		var s;
		var l = e.match(sl);
		if (l) ul(l[0], i, n, t);
		if ((s = e.match(/drawing r:id="(.*?)"/))) i["!rel"] = s[1];
		if (a["!id"][i["!rel"]]) i["!drawel"] = a["!id"][i["!rel"]];
		return i
	}
	var Il = [["allowRefreshQuery", false, "bool"], ["autoCompressPictures", true, "bool"], ["backupFile", false, "bool"], ["checkCompatibility", false, "bool"], ["CodeName", ""], ["date1904", false, "bool"], ["defaultThemeVersion", 0, "int"], ["filterPrivacy", false, "bool"], ["hidePivotFieldList", false, "bool"], ["promptedSolutions", false, "bool"], ["publishItems", false, "bool"], ["refreshAllConnections", false, "bool"], ["saveExternalLinkValues", true, "bool"], ["showBorderUnselectedTables", true, "bool"], ["showInkAnnotation", true, "bool"], ["showObjects", "all"], ["showPivotChartFilter", false, "bool"], ["updateLinks", "userSet"], ];
	var Pl = [["activeTab", 0, "int"], ["autoFilterDateGrouping", true, "bool"], ["firstSheet", 0, "int"], ["minimized", false, "bool"], ["showHorizontalScroll", true, "bool"], ["showSheetTabs", true, "bool"], ["showVerticalScroll", true, "bool"], ["tabRatio", 600, "int"], ["visibility", "visible"], ];
	var Rl = [];
	var Ll = [["calcCompleted", "true"], ["calcMode", "auto"], ["calcOnSave", "true"], ["concurrentCalc", "true"], ["fullCalcOnLoad", "false"], ["fullPrecision", "true"], ["iterate", "false"], ["iterateCount", "100"], ["iterateDelta", "0.001"], ["refMode", "A1"], ];
	function Bl(e, r) {
		for (var t = 0; t != e.length; ++t) {
			var a = e[t];
			for (var n = 0; n != r.length; ++n) {
				var i = r[n];
				if (a[i[0]] == null) a[i[0]] = i[1];
				else switch (i[2]) {
				case "bool":
					if (typeof a[i[0]] == "string") a[i[0]] = mt(a[i[0]]);
					break;
				case "int":
					if (typeof a[i[0]] == "string") a[i[0]] = parseInt(a[i[0]], 10);
					break
				}
			}
		}
	}
	function Ul(e, r) {
		for (var t = 0; t != r.length; ++t) {
			var a = r[t];
			if (e[a[0]] == null) e[a[0]] = a[1];
			else switch (a[2]) {
			case "bool":
				if (typeof e[a[0]] == "string") e[a[0]] = mt(e[a[0]]);
				break;
			case "int":
				if (typeof e[a[0]] == "string") e[a[0]] = parseInt(e[a[0]], 10);
				break
			}
		}
	}
	function zl(e) {
		Ul(e.WBProps, Il);
		Ul(e.CalcPr, Ll);
		Bl(e.WBView, Pl);
		Bl(e.Sheets, Rl);
		Hs.date1904 = mt(e.WBProps.date1904)
	}
	var Wl = ":][*?/\\".split("");
	function jl(e, r) {
		try {
			if (e == "") throw new Error("Sheet name cannot be blank");
			if (e.length > 31) throw new Error("Sheet name cannot exceed 31 chars");
			if (e.charCodeAt(0) == 39 || e.charCodeAt(e.length - 1) == 39) throw new Error("Sheet name cannot start or end with apostrophe (')");
			if (e.toLowerCase() == "history") throw new Error("Sheet name cannot be 'History'");
			Wl.forEach(function(r) {
				if (e.indexOf(r) == -1) return;
				throw new Error("Sheet name cannot contain : \\ / ? * [ ]");
			})
		} catch(t) {
			if (r) return false;
			throw t;
		}
		return true
	}
	var Xl = /<\w+:workbook/;
	function Gl(e, r) {
		if (!e) throw new Error("Could not find file");
		var t = {
			AppVersion: {},
			WBProps: {},
			WBView: [],
			Sheets: [],
			CalcPr: {},
			Names: [],
			xmlns: "",
		};
		var a = false,
		n = "xmlns";
		var i = {},
		s = 0;
		e.replace(qr,
		function l(o, c) {
			var f = rt(o);
			switch (tt(f[0])) {
			case "<?xml":
				break;
			case "<workbook":
				if (o.match(Xl)) n = "xmlns" + o.match(/<(\w+):/)[1];
				t.xmlns = f[n];
				break;
			case "</workbook>":
				break;
			case "<fileVersion":
				delete f[0];
				t.AppVersion = f;
				break;
			case "<fileVersion/>":
			case "</fileVersion>":
				break;
			case "<fileSharing":
				break;
			case "<fileSharing/>":
				break;
			case "<workbookPr":
			case "<workbookPr/>":
				Il.forEach(function(e) {
					if (f[e[0]] == null) return;
					switch (e[2]) {
					case "bool":
						t.WBProps[e[0]] = mt(f[e[0]]);
						break;
					case "int":
						t.WBProps[e[0]] = parseInt(f[e[0]], 10);
						break;
					default:
						t.WBProps[e[0]] = f[e[0]]
					}
				});
				if (f.codeName) t.WBProps.CodeName = kt(f.codeName);
				break;
			case "</workbookPr>":
				break;
			case "<workbookProtection":
				break;
			case "<workbookProtection/>":
				break;
			case "<bookViews":
			case "<bookViews>":
			case "</bookViews>":
				break;
			case "<workbookView":
			case "<workbookView/>":
				delete f[0];
				t.WBView.push(f);
				break;
			case "</workbookView>":
				break;
			case "<sheets":
			case "<sheets>":
			case "</sheets>":
				break;
			case "<sheet":
				switch (f.state) {
				case "hidden":
					f.Hidden = 1;
					break;
				case "veryHidden":
					f.Hidden = 2;
					break;
				default:
					f.Hidden = 0
				}
				delete f.state;
				f.name = it(kt(f.name));
				delete f[0];
				t.Sheets.push(f);
				break;
			case "</sheet>":
				break;
			case "<functionGroups":
			case "<functionGroups/>":
				break;
			case "<functionGroup":
				break;
			case "<externalReferences":
			case "</externalReferences>":
			case "<externalReferences>":
				break;
			case "<externalReference":
				break;
			case "<definedNames/>":
				break;
			case "<definedNames>":
			case "<definedNames":
				a = true;
				break;
			case "</definedNames>":
				a = false;
				break;
			case "<definedName":
				{
					i = {};
					i.Name = kt(f.name);
					if (f.comment) i.Comment = f.comment;
					if (f.localSheetId) i.Sheet = +f.localSheetId;
					if (mt(f.hidden || "0")) i.Hidden = true;
					s = c + o.length
				}
				break;
			case "</definedName>":
				{
					i.Ref = it(kt(e.slice(s, c)));
					t.Names.push(i)
				}
				break;
			case "<definedName/>":
				break;
			case "<calcPr":
				delete f[0];
				t.CalcPr = f;
				break;
			case "<calcPr/>":
				delete f[0];
				t.CalcPr = f;
				break;
			case "</calcPr>":
				break;
			case "<oleSize":
				break;
			case "<customWorkbookViews>":
			case "</customWorkbookViews>":
			case "<customWorkbookViews":
				break;
			case "<customWorkbookView":
			case "</customWorkbookView>":
				break;
			case "<pivotCaches>":
			case "</pivotCaches>":
			case "<pivotCaches":
				break;
			case "<pivotCache":
				break;
			case "<smartTagPr":
			case "<smartTagPr/>":
				break;
			case "<smartTagTypes":
			case "<smartTagTypes>":
			case "</smartTagTypes>":
				break;
			case "<smartTagType":
				break;
			case "<webPublishing":
			case "<webPublishing/>":
				break;
			case "<fileRecoveryPr":
			case "<fileRecoveryPr/>":
				break;
			case "<webPublishObjects>":
			case "<webPublishObjects":
			case "</webPublishObjects>":
				break;
			case "<webPublishObject":
				break;
			case "<extLst":
			case "<extLst>":
			case "</extLst>":
			case "<extLst/>":
				break;
			case "<ext":
				a = true;
				break;
			case "</ext>":
				a = false;
				break;
			case "<ArchID":
				break;
			case "<AlternateContent":
			case "<AlternateContent>":
				a = true;
				break;
			case "</AlternateContent>":
				a = false;
				break;
			case "<revisionPtr":
				break;
			default:
				if (!a && r.WTF) throw new Error("unrecognized " + f[0] + " in workbook");
			}
			return o
		});
		if (Lt.indexOf(t.xmlns) === -1) throw new Error("Unknown Namespace: " + t.xmlns);
		zl(t);
		return t
	}
	function Jl(e, r, t) {
		if (r.slice(-4) === ".bin") return parse_wb_bin(e, t);
		return Gl(e, t)
	}
	function Zl(e, r, t, a, n, i, s, l) {
		if (r.slice(-4) === ".bin") return parse_ws_bin(e, a, t, n, i, s, l);
		return cl(e, a, t, n, i, s, l)
	}
	function Kl(e, r, t, a, n, i, s, l) {
		if (r.slice(-4) === ".bin") return parse_cs_bin(e, a, t, n, i, s, l);
		return Nl(e, a, t, n, i, s, l)
	}
	function ql(e, r, t, a, n, i, s, l) {
		if (r.slice(-4) === ".bin") return Fs(e, a, t, n, i, s, l);
		return Ds(e, a, t, n, i, s, l)
	}
	function Ql(e, r, t, a, n, i, s, l) {
		if (r.slice(-4) === ".bin") return Ts(e, a, t, n, i, s, l);
		return Es(e, a, t, n, i, s, l)
	}
	function eo(e, r, t, a) {
		if (r.slice(-4) === ".bin") return parse_sty_bin(e, t, a);
		return Yi(e, t, a)
	}
	function ro(e, r, t) {
		if (r.slice(-4) === ".bin") return parse_sst_bin(e, t);
		return vi(e, t)
	}
	function to(e, r, t) {
		if (r.slice(-4) === ".bin") return parse_comments_bin(e, t);
		return gs(e, t)
	}
	function ao(e, r, t) {
		if (r.slice(-4) === ".bin") return parse_cc_bin(e, r, t);
		return parse_cc_xml(e, r, t)
	}
	function no(e, r, t, a) {
		if (t.slice(-4) === ".bin") return fs(e, r, t, a);
		return cs(e, r, t, a)
	}
	function io(e, r, t) {
		if (r.slice(-4) === ".bin") return parse_xlmeta_bin(e, r, t);
		return ls(e, r, t)
	}
	function lo(e, r, t, a) {
		var n = e["!merges"] || [];
		var i = [];
		var s = {};
		var l = e["!data"] != null;
		for (var o = r.s.c; o <= r.e.c; ++o) {
			var c = 0,
			f = 0;
			for (var u = 0; u < n.length; ++u) {
				if (n[u].s.r > t || n[u].s.c > o) continue;
				if (n[u].e.r < t || n[u].e.c < o) continue;
				if (n[u].s.r < t || n[u].s.c < o) {
					c = -1;
					break
				}
				c = n[u].e.r - n[u].s.r + 1;
				f = n[u].e.c - n[u].s.c + 1;
				break
			}
			if (c < 0) continue;
			var h = Ra(o) + Ma(t);
			var d = l ? (e["!data"][t] || [])[o] : e[h];
			var p = (d && d.v != null && (d.h || ut(d.w || (Ya(d), d.w) || ""))) || "";
			s = {};
			if (c > 1) s.rowspan = c;
			if (f > 1) s.colspan = f;
			if (a.editable) p = '<span contenteditable="true">' + p + "</span>";
			else if (d) {
				s["data-t"] = (d && d.t) || "z";
				if (d.v != null) s["data-v"] = d.v instanceof Date ? d.v.toISOString() : d.v;
				if (d.z != null) s["data-z"] = d.z;
				if (d.l && (d.l.Target || "#").charAt(0) != "#") p = '<a href="' + ut(d.l.Target) + '">' + p + "</a>"
			}
			s.id = (a.id || "sjs") + "-" + h;
			i.push(Ot("td", p, s))
		}
		var m = "<tr>";
		return m + i.join("") + "</tr>"
	}
	var oo = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
	var co = "</body></html>";
	function uo(e, r, t) {
		var a = [];
		return (a.join("") + "<table" + (t && t.id ? ' id="' + t.id + '"': "") + ">")
	}
	function ho(e, r) {
		var t = r || {};
		var a = t.header != null ? t.header: oo;
		var n = t.footer != null ? t.footer: co;
		var i = [a];
		var s = Wa(e["!ref"] || "A1");
		i.push(uo(e, s, t));
		if (e["!ref"]) for (var l = s.s.r; l <= s.e.r; ++l) i.push(lo(e, s, l, t));
		i.push("</table>" + n);
		return i.join("")
	}
	function po(e, r, t) {
		var a = r.rows;
		if (!a) {
			throw "Unsupported origin when " + r.tagName + " is not a TABLE";
		}
		var n = t || {};
		var i = e["!data"] != null;
		var s = 0,
		l = 0;
		if (n.origin != null) {
			if (typeof n.origin == "number") s = n.origin;
			else {
				var o = typeof n.origin == "string" ? za(n.origin) : n.origin;
				s = o.r;
				l = o.c
			}
		}
		var c = Math.min(n.sheetRows || 1e7, a.length);
		var f = {
			s: {
				r: 0,
				c: 0,
			},
			e: {
				r: s,
				c: l,
			},
		};
		if (e["!ref"]) {
			var u = Wa(e["!ref"]);
			f.s.r = Math.min(f.s.r, u.s.r);
			f.s.c = Math.min(f.s.c, u.s.c);
			f.e.r = Math.max(f.e.r, u.e.r);
			f.e.c = Math.max(f.e.c, u.e.c);
			if (s == -1) f.e.r = s = u.e.r + 1
		}
		var h = [],
		d = 0;
		var p = e["!rows"] || (e["!rows"] = []);
		var m = 0,
		v = 0,
		g = 0,
		b = 0,
		w = 0,
		k = 0;
		if (!e["!cols"]) e["!cols"] = [];
		for (; m < a.length && v < c; ++m) {
			var y = a[m];
			if (go(y)) {
				if (n.display) continue;
				p[v] = {
					hidden: true,
				}
			}
			var x = y.cells;
			for (g = b = 0; g < x.length; ++g) {
				var S = x[g];
				if (n.display && go(S)) continue;
				var C = S.hasAttribute("data-v") ? S.getAttribute("data-v") : S.hasAttribute("v") ? S.getAttribute("v") : St(S.innerHTML);
				var _ = S.getAttribute("data-z") || S.getAttribute("z");
				for (d = 0; d < h.length; ++d) {
					var A = h[d];
					if (A.s.c == b + l && A.s.r < v + s && v + s <= A.e.r) {
						b = A.e.c + 1 - l;
						d = -1
					}
				}
				k = +S.getAttribute("colspan") || 1;
				if ((w = +S.getAttribute("rowspan") || 1) > 1 || k > 1) h.push({
					s: {
						r: v + s,
						c: b + l,
					},
					e: {
						r: v + s + (w || 1) - 1,
						c: b + l + (k || 1) - 1,
					},
				});
				var T = {
					t: "s",
					v: C,
				};
				var E = S.getAttribute("data-t") || S.getAttribute("t") || "";
				if (C != null) {
					if (C.length == 0) T.t = E || "z";
					else if (n.raw || C.trim().length == 0 || E == "s") {} else if (C === "TRUE") T = {
						t: "b",
						v: true,
					};
					else if (C === "FALSE") T = {
						t: "b",
						v: false,
					};
					else if (!isNaN(Sr(C))) T = {
						t: "n",
						v: Sr(C),
					};
					else if (!isNaN(Or(C).getDate())) {
						T = {
							t: "d",
							v: wr(C),
						};
						if (n.UTC) T.v = Ir(T.v);
						if (!n.cellDates) T = {
							t: "n",
							v: dr(T.v),
						};
						T.z = n.dateNF || q[14]
					}
				}
				if (T.z === undefined && _ != null) T.z = _;
				var F = "",
				D = S.getElementsByTagName("A");
				if (D && D.length) for (var O = 0; O < D.length; ++O) if (D[O].hasAttribute("href")) {
					F = D[O].getAttribute("href");
					if (F.charAt(0) != "#") break
				}
				if (F && F.charAt(0) != "#" && F.slice(0, 11).toLowerCase() != "javascript:") T.l = {
					Target: F,
				};
				if (i) {
					if (!e["!data"][v + s]) e["!data"][v + s] = [];
					e["!data"][v + s][b + l] = T
				} else e[$a({
					c: b + l,
					r: v + s,
				})] = T;
				if (f.e.c < b + l) f.e.c = b + l;
				b += k
			}++v
		}
		if (h.length) e["!merges"] = (e["!merges"] || []).concat(h);
		f.e.r = Math.max(f.e.r, v - 1 + s);
		e["!ref"] = ja(f);
		if (v >= c) e["!fullref"] = ja(((f.e.r = a.length - m + v - 1 + s), f));
		return e
	}
	function mo(e, r) {
		var t = r || {};
		var a = {};
		if (t.dense) a["!data"] = [];
		return po(a, e, r)
	}
	function vo(e, r) {
		var t = Ja(mo(e, r), r);
		return t
	}
	function go(e) {
		var r = "";
		var t = bo(e);
		if (t) r = t(e).getPropertyValue("display");
		if (!r) r = e.style && e.style.display;
		return r === "none"
	}
	function bo(e) {
		if (e.ownerDocument.defaultView && typeof e.ownerDocument.defaultView.getComputedStyle === "function") return e.ownerDocument.defaultView.getComputedStyle;
		if (typeof getComputedStyle === "function") return getComputedStyle;
		return null
	}
	function wo(e) {
		var r = e.replace(/[\t\r\n]/g, " ").trim().replace(/ +/g, " ").replace(/<text:s\/>/g, " ").replace(/<text:s text:c="(\d+)"\/>/g,
		function(e, r) {
			return Array(parseInt(r, 10) + 1).join(" ")
		}).replace(/<text:tab[^>]*\/>/g, "\t").replace(/<text:line-break\/>/g, "\n");
		var t = it(r.replace(/<[^>]*>/g, ""));
		return [t]
	}
	function ko(e, r, t) {
		var a = t || {};
		var n = It(e);
		Pt.lastIndex = 0;
		n = n.replace(/<!--([\s\S]*?)-->/gm, "").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm, "");
		var i, s, l = "",
		o = "",
		c, f = 0,
		u = -1,
		h = false,
		d = "";
		while ((i = Pt.exec(n))) {
			switch ((i[3] = i[3].replace(/_.*$/, ""))) {
			case "number-style":
			case "currency-style":
			case "percentage-style":
			case "date-style":
			case "time-style":
			case "text-style":
				if (i[1] === "/") {
					h = false;
					if (s["truncate-on-overflow"] == "false") {
						if (l.match(/h/)) l = l.replace(/h+/, "[$&]");
						else if (l.match(/m/)) l = l.replace(/m+/, "[$&]");
						else if (l.match(/s/)) l = l.replace(/s+/, "[$&]")
					}
					a[s.name] = l;
					l = ""
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					h = true;
					l = "";
					s = rt(i[0], false)
				}
				break;
			case "boolean-style":
				if (i[1] === "/") {
					h = false;
					a[s.name] = "General";
					l = ""
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					h = true;
					l = "";
					s = rt(i[0], false)
				}
				break;
			case "boolean":
				l += "General";
				break;
			case "text":
				if (i[1] === "/") {
					d = n.slice(u, Pt.lastIndex - i[0].length);
					if (d == "%" && s[0] == "<number:percentage-style") l += "%";
					else l += '"' + d.replace(/"/g, '""') + '"'
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					u = Pt.lastIndex
				}
				break;
			case "day":
				{
					c = rt(i[0], false);
					switch (c["style"]) {
					case "short":
						l += "d";
						break;
					case "long":
						l += "dd";
						break;
					default:
						l += "dd";
						break
					}
				}
				break;
			case "day-of-week":
				{
					c = rt(i[0], false);
					switch (c["style"]) {
					case "short":
						l += "ddd";
						break;
					case "long":
						l += "dddd";
						break;
					default:
						l += "ddd";
						break
					}
				}
				break;
			case "era":
				{
					c = rt(i[0], false);
					switch (c["style"]) {
					case "short":
						l += "ee";
						break;
					case "long":
						l += "eeee";
						break;
					default:
						l += "eeee";
						break
					}
				}
				break;
			case "hours":
				{
					c = rt(i[0], false);
					switch (c["style"]) {
					case "short":
						l += "h";
						break;
					case "long":
						l += "hh";
						break;
					default:
						l += "hh";
						break
					}
				}
				break;
			case "minutes":
				{
					c = rt(i[0], false);
					switch (c["style"]) {
					case "short":
						l += "m";
						break;
					case "long":
						l += "mm";
						break;
					default:
						l += "mm";
						break
					}
				}
				break;
			case "month":
				{
					c = rt(i[0], false);
					if (c["textual"]) l += "mm";
					switch (c["style"]) {
					case "short":
						l += "m";
						break;
					case "long":
						l += "mm";
						break;
					default:
						l += "m";
						break
					}
				}
				break;
			case "seconds":
				{
					c = rt(i[0], false);
					switch (c["style"]) {
					case "short":
						l += "s";
						break;
					case "long":
						l += "ss";
						break;
					default:
						l += "ss";
						break
					}
					if (c["decimal-places"]) l += "." + xr("0", +c["decimal-places"])
				}
				break;
			case "year":
				{
					c = rt(i[0], false);
					switch (c["style"]) {
					case "short":
						l += "yy";
						break;
					case "long":
						l += "yyyy";
						break;
					default:
						l += "yy";
						break
					}
				}
				break;
			case "am-pm":
				l += "AM/PM";
				break;
			case "week-of-year":
			case "quarter":
				console.error("Excel does not support ODS format token " + i[3]);
				break;
			case "fill-character":
				if (i[1] === "/") {
					d = n.slice(u, Pt.lastIndex - i[0].length);
					l += '"' + d.replace(/"/g, '""') + '"*'
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					u = Pt.lastIndex
				}
				break;
			case "scientific-number":
				c = rt(i[0], false);
				l += "0." + xr("0", +c["min-decimal-places"] || +c["decimal-places"] || 2) + xr("?", +c["decimal-places"] - +c["min-decimal-places"] || 0) + "E" + (mt(c["forced-exponent-sign"]) ? "+": "") + xr("0", +c["min-exponent-digits"] || 2);
				break;
			case "fraction":
				c = rt(i[0], false);
				if (!+c["min-integer-digits"]) l += "#";
				else l += xr("0", +c["min-integer-digits"]);
				l += " ";
				l += xr("?", +c["min-numerator-digits"] || 1);
				l += "/";
				if ( + c["denominator-value"]) l += c["denominator-value"];
				else l += xr("?", +c["min-denominator-digits"] || 1);
				break;
			case "currency-symbol":
				if (i[1] === "/") {
					l += '"' + n.slice(u, Pt.lastIndex - i[0].length).replace(/"/g, '""') + '"'
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					u = Pt.lastIndex
				} else l += "$";
				break;
			case "text-properties":
				c = rt(i[0], false);
				switch ((c["color"] || "").toLowerCase().replace("#", "")) {
				case "ff0000":
				case "red":
					l = "[Red]" + l;
					break
				}
				break;
			case "text-content":
				l += "@";
				break;
			case "map":
				c = rt(i[0], false);
				if (it(c["condition"]) == "value()>=0") l = a[c["apply-style-name"]] + ";" + l;
				else console.error("ODS number format may be incorrect: " + c["condition"]);
				break;
			case "number":
				if (i[1] === "/") break;
				c = rt(i[0], false);
				o = "";
				o += xr("0", +c["min-integer-digits"] || 1);
				if (mt(c["grouping"])) o = he(xr("#", Math.max(0, 4 - o.length)) + o);
				if ( + c["min-decimal-places"] || +c["decimal-places"]) o += ".";
				if ( + c["min-decimal-places"]) o += xr("0", +c["min-decimal-places"] || 1);
				if ( + c["decimal-places"] - ( + c["min-decimal-places"] || 0)) o += xr("0", +c["decimal-places"] - ( + c["min-decimal-places"] || 0));
				l += o;
				break;
			case "embedded-text":
				if (i[1] === "/") {
					if (f == 0) l += '"' + n.slice(u, Pt.lastIndex - i[0].length).replace(/"/g, '""') + '"';
					else l = l.slice(0, f) + '"' + n.slice(u, Pt.lastIndex - i[0].length).replace(/"/g, '""') + '"' + l.slice(f)
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					u = Pt.lastIndex;
					f = -+rt(i[0], false)["position"] || 0
				}
				break
			}
		}
		return a
	}
	function yo(e, r, t) {
		var a = r || {};
		if (b != null && a.dense == null) a.dense = b;
		var n = It(e);
		var i = [],
		s;
		var l;
		var o, c = "",
		f = 0;
		var u;
		var h;
		var d = {},
		p = [];
		var m = {};
		if (a.dense) m["!data"] = [];
		var v, g;
		var w = {
			value: "",
		};
		var k = "",
		y = 0,
		x, S = "",
		C = 0;
		var _ = [],
		A = [];
		var T = -1,
		E = -1,
		F = {
			s: {
				r: 1e6,
				c: 1e7,
			},
			e: {
				r: 0,
				c: 0,
			},
		};
		var D = 0;
		var O = t || {},
		M = {};
		var N = [],
		I = {},
		P = 0,
		R = 0;
		var L = [],
		B = 1,
		U = 1;
		var z = [];
		var $ = {
			Names: [],
			WBProps: {},
		};
		var W = {};
		var j = ["", ""];
		var H = [],
		V = {};
		var X = "",
		G = 0;
		var Y = false,
		J = false;
		var Z = 0;
		Pt.lastIndex = 0;
		n = n.replace(/<!--([\s\S]*?)-->/gm, "").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm, "");
		while ((v = Pt.exec(n))) switch ((v[3] = v[3].replace(/_.*$/, ""))) {
		case "table":
		case "工作表":
			if (v[1] === "/") {
				if (F.e.c >= F.s.c && F.e.r >= F.s.r) m["!ref"] = ja(F);
				else m["!ref"] = "A1:A1";
				if (a.sheetRows > 0 && a.sheetRows <= F.e.r) {
					m["!fullref"] = m["!ref"];
					F.e.r = a.sheetRows - 1;
					m["!ref"] = ja(F)
				}
				if (N.length) m["!merges"] = N;
				if (L.length) m["!rows"] = L;
				u.name = u["名称"] || u.name;
				if (typeof JSON !== "undefined") JSON.stringify(u);
				p.push(u.name);
				d[u.name] = m;
				J = false
			} else if (v[0].charAt(v[0].length - 2) !== "/") {
				u = rt(v[0], false);
				T = E = -1;
				F.s.r = F.s.c = 1e7;
				F.e.r = F.e.c = 0;
				m = {};
				if (a.dense) m["!data"] = [];
				N = [];
				L = [];
				J = true
			}
			break;
		case "table-row-group":
			if (v[1] === "/")--D;
			else++D;
			break;
		case "table-row":
		case "行":
			if (v[1] === "/") {
				T += B;
				B = 1;
				break
			}
			h = rt(v[0], false);
			if (h["行号"]) T = h["行号"] - 1;
			else if (T == -1) T = 0;
			B = +h["number-rows-repeated"] || 1;
			if (B < 10) for (Z = 0; Z < B; ++Z) if (D > 0) L[T + Z] = {
				level: D,
			};
			E = -1;
			break;
		case "covered-table-cell":
			if (v[1] !== "/")++E;
			if (a.sheetStubs) {
				if (a.dense) {
					if (!m["!data"][T]) m["!data"][T] = [];
					m["!data"][T][E] = {
						t: "z",
					}
				} else m[$a({
					r: T,
					c: E,
				})] = {
					t: "z",
				}
			}
			k = "";
			_ = [];
			break;
		case "table-cell":
		case "数据":
			if (v[0].charAt(v[0].length - 2) === "/") {++E;
				w = rt(v[0], false);
				U = parseInt(w["number-columns-repeated"] || "1", 10);
				g = {
					t: "z",
					v: null,
				};
				if (w.formula && a.cellFormula != false) g.f = Us(it(w.formula));
				if (w["style-name"] && M[w["style-name"]]) g.z = M[w["style-name"]];
				if ((w["数据类型"] || w["value-type"]) == "string") {
					g.t = "s";
					g.v = it(w["string-value"] || "");
					if (a.dense) {
						if (!m["!data"][T]) m["!data"][T] = [];
						m["!data"][T][E] = g
					} else {
						m[Ra(E) + Ma(T)] = g
					}
				}
				E += U - 1
			} else if (v[1] !== "/") {++E;
				k = S = "";
				y = C = 0;
				_ = [];
				A = [];
				U = 1;
				var K = B ? T + B - 1 : T;
				if (E > F.e.c) F.e.c = E;
				if (E < F.s.c) F.s.c = E;
				if (T < F.s.r) F.s.r = T;
				if (K > F.e.r) F.e.r = K;
				w = rt(v[0], false);
				H = [];
				V = {};
				g = {
					t: w["数据类型"] || w["value-type"],
					v: null,
				};
				if (w["style-name"] && M[w["style-name"]]) g.z = M[w["style-name"]];
				if (a.cellFormula) {
					if (w.formula) w.formula = it(w.formula);
					if (w["number-matrix-columns-spanned"] && w["number-matrix-rows-spanned"]) {
						P = parseInt(w["number-matrix-rows-spanned"], 10) || 0;
						R = parseInt(w["number-matrix-columns-spanned"], 10) || 0;
						I = {
							s: {
								r: T,
								c: E,
							},
							e: {
								r: T + P - 1,
								c: E + R - 1,
							},
						};
						g.F = ja(I);
						z.push([I, g.F])
					}
					if (w.formula) g.f = Us(w.formula);
					else for (Z = 0; Z < z.length; ++Z) if (T >= z[Z][0].s.r && T <= z[Z][0].e.r) if (E >= z[Z][0].s.c && E <= z[Z][0].e.c) g.F = z[Z][1]
				}
				if (w["number-columns-spanned"] || w["number-rows-spanned"]) {
					P = parseInt(w["number-rows-spanned"], 10) || 0;
					R = parseInt(w["number-columns-spanned"], 10) || 0;
					I = {
						s: {
							r: T,
							c: E,
						},
						e: {
							r: T + P - 1,
							c: E + R - 1,
						},
					};
					N.push(I)
				}
				if (w["number-columns-repeated"]) U = parseInt(w["number-columns-repeated"], 10);
				switch (g.t) {
				case "boolean":
					g.t = "b";
					g.v = mt(w["boolean-value"]) || +w["boolean-value"] >= 1;
					break;
				case "float":
					g.t = "n";
					g.v = parseFloat(w.value);
					if (a.cellDates && g.z && Re(g.z)) {
						g.v = pr(g.v + ($.WBProps.date1904 ? 1462 : 0));
						g.t = typeof g.v == "number" ? "n": "d"
					}
					break;
				case "percentage":
					g.t = "n";
					g.v = parseFloat(w.value);
					break;
				case "currency":
					g.t = "n";
					g.v = parseFloat(w.value);
					break;
				case "date":
					g.t = "d";
					g.v = wr(w["date-value"], $.WBProps.date1904);
					if (!a.cellDates) {
						g.t = "n";
						g.v = dr(g.v, $.WBProps.date1904)
					}
					if (!g.z) g.z = "m/d/yy";
					break;
				case "time":
					g.t = "n";
					g.v = mr(w["time-value"]) / 86400;
					if (a.cellDates) {
						g.v = pr(g.v);
						g.t = typeof g.v == "number" ? "n": "d"
					}
					if (!g.z) g.z = "HH:MM:SS";
					break;
				case "number":
					g.t = "n";
					g.v = parseFloat(w["数据数值"]);
					break;
				default:
					if (g.t === "string" || g.t === "text" || !g.t) {
						g.t = "s";
						if (w["string-value"] != null) {
							k = it(w["string-value"]);
							_ = []
						}
					} else throw new Error("Unsupported value type " + g.t);
				}
			} else {
				Y = false;
				if (g.t === "s") {
					g.v = k || "";
					if (_.length) g.R = _;
					Y = y == 0
				}
				if (W.Target) g.l = W;
				if (H.length > 0) {
					g.c = H;
					H = []
				}
				if (k && a.cellText !== false) g.w = k;
				if (Y) {
					g.t = "z";
					delete g.v
				}
				if (!Y || a.sheetStubs) {
					if (! (a.sheetRows && a.sheetRows <= T)) {
						for (var q = 0; q < B; ++q) {
							U = parseInt(w["number-columns-repeated"] || "1", 10);
							if (a.dense) {
								if (!m["!data"][T + q]) m["!data"][T + q] = [];
								m["!data"][T + q][E] = q == 0 ? g: yr(g);
								while (--U > 0) m["!data"][T + q][E + U] = yr(g)
							} else {
								m[$a({
									r: T + q,
									c: E,
								})] = g;
								while (--U > 0) m[$a({
									r: T + q,
									c: E + U,
								})] = yr(g)
							}
							if (F.e.c <= E) F.e.c = E
						}
					}
				}
				U = parseInt(w["number-columns-repeated"] || "1", 10);
				E += U - 1;
				U = 0;
				g = {};
				k = "";
				_ = []
			}
			W = {};
			break;
		case "document":
		case "document-content":
		case "电子表格文档":
		case "spreadsheet":
		case "主体":
		case "scripts":
		case "styles":
		case "font-face-decls":
		case "master-styles":
			if (v[1] === "/") {
				if ((s = i.pop())[0] !== v[3]) throw "Bad state: " + s;
			} else if (v[0].charAt(v[0].length - 2) !== "/") i.push([v[3], true]);
			break;
		case "annotation":
			if (v[1] === "/") {
				if ((s = i.pop())[0] !== v[3]) throw "Bad state: " + s;
				V.t = k;
				if (_.length) V.R = _;
				V.a = X;
				H.push(V);
				k = S;
				y = C;
				_ = A
			} else if (v[0].charAt(v[0].length - 2) !== "/") {
				i.push([v[3], false]);
				var Q = rt(v[0], true);
				if (! (Q["display"] && mt(Q["display"]))) H.hidden = true;
				S = k;
				C = y;
				A = _;
				k = "";
				y = 0;
				_ = []
			}
			X = "";
			G = 0;
			break;
		case "creator":
			if (v[1] === "/") {
				X = n.slice(G, v.index)
			} else G = v.index + v[0].length;
			break;
		case "meta":
		case "元数据":
		case "settings":
		case "config-item-set":
		case "config-item-map-indexed":
		case "config-item-map-entry":
		case "config-item-map-named":
		case "shapes":
		case "frame":
		case "text-box":
		case "image":
		case "data-pilot-tables":
		case "list-style":
		case "form":
		case "dde-links":
		case "event-listeners":
		case "chart":
			if (v[1] === "/") {
				if ((s = i.pop())[0] !== v[3]) throw "Bad state: " + s;
			} else if (v[0].charAt(v[0].length - 2) !== "/") i.push([v[3], false]);
			k = "";
			y = 0;
			_ = [];
			break;
		case "scientific-number":
		case "currency-symbol":
		case "fill-character":
			break;
		case "text-style":
		case "boolean-style":
		case "number-style":
		case "currency-style":
		case "percentage-style":
		case "date-style":
		case "time-style":
			if (v[1] === "/") {
				var ee = Pt.lastIndex;
				ko(n.slice(o, Pt.lastIndex), r, O);
				Pt.lastIndex = ee
			} else if (v[0].charAt(v[0].length - 2) !== "/") {
				o = Pt.lastIndex - v[0].length
			}
			break;
		case "script":
			break;
		case "libraries":
			break;
		case "automatic-styles":
			break;
		case "default-style":
		case "page-layout":
			break;
		case "style":
			{
				var re = rt(v[0], false);
				if (re["family"] == "table-cell" && O[re["data-style-name"]]) M[re["name"]] = O[re["data-style-name"]]
			}
			break;
		case "map":
			break;
		case "font-face":
			break;
		case "paragraph-properties":
			break;
		case "table-properties":
			break;
		case "table-column-properties":
			break;
		case "table-row-properties":
			break;
		case "table-cell-properties":
			break;
		case "number":
			break;
		case "fraction":
			break;
		case "day":
		case "month":
		case "year":
		case "era":
		case "day-of-week":
		case "week-of-year":
		case "quarter":
		case "hours":
		case "minutes":
		case "seconds":
		case "am-pm":
			break;
		case "boolean":
			break;
		case "text":
			if (v[0].slice(-2) === "/>") break;
			else if (v[1] === "/") switch (i[i.length - 1][0]) {
			case "number-style":
			case "date-style":
			case "time-style":
				c += n.slice(f, v.index);
				break
			} else f = v.index + v[0].length;
			break;
		case "named-range":
			l = rt(v[0], false);
			j = $s(l["cell-range-address"]);
			var te = {
				Name: l.name,
				Ref: j[0] + "!" + j[1],
			};
			if (J) te.Sheet = p.length;
			$.Names.push(te);
			break;
		case "text-content":
			break;
		case "text-properties":
			break;
		case "embedded-text":
			break;
		case "body":
		case "电子表格":
			break;
		case "forms":
			break;
		case "table-column":
			break;
		case "table-header-rows":
			break;
		case "table-rows":
			break;
		case "table-column-group":
			break;
		case "table-header-columns":
			break;
		case "table-columns":
			break;
		case "null-date":
			l = rt(v[0], false);
			switch (l["date-value"]) {
			case "1904-01-01":
				$.WBProps.date1904 = true;
				break
			}
			break;
		case "graphic-properties":
			break;
		case "calculation-settings":
			break;
		case "named-expressions":
			break;
		case "label-range":
			break;
		case "label-ranges":
			break;
		case "named-expression":
			break;
		case "sort":
			break;
		case "sort-by":
			break;
		case "sort-groups":
			break;
		case "tab":
			break;
		case "line-break":
			break;
		case "span":
			break;
		case "p":
		case "文本串":
			if (["master-styles"].indexOf(i[i.length - 1][0]) > -1) break;
			if (v[1] === "/" && (!w || !w["string-value"])) {
				var ae = wo(n.slice(y, v.index), x);
				k = (k.length > 0 ? k + "\n": "") + ae[0]
			} else if (v[0].slice(-2) == "/>") {
				k += "\n"
			} else {
				x = rt(v[0], false);
				y = v.index + v[0].length
			}
			break;
		case "s":
			break;
		case "database-range":
			if (v[1] === "/") break;
			try {
				j = $s(rt(v[0])["target-range-address"]);
				d[j[0]]["!autofilter"] = {
					ref: j[1],
				}
			} catch(ne) {}
			break;
		case "date":
			break;
		case "object":
			break;
		case "title":
		case "标题":
			break;
		case "desc":
			break;
		case "binary-data":
			break;
		case "table-source":
			break;
		case "scenario":
			break;
		case "iteration":
			break;
		case "content-validations":
			break;
		case "content-validation":
			break;
		case "help-message":
			break;
		case "error-message":
			break;
		case "database-ranges":
			break;
		case "filter":
			break;
		case "filter-and":
			break;
		case "filter-or":
			break;
		case "filter-condition":
			break;
		case "filter-set-item":
			break;
		case "list-level-style-bullet":
			break;
		case "list-level-style-number":
			break;
		case "list-level-properties":
			break;
		case "sender-firstname":
		case "sender-lastname":
		case "sender-initials":
		case "sender-title":
		case "sender-position":
		case "sender-email":
		case "sender-phone-private":
		case "sender-fax":
		case "sender-company":
		case "sender-phone-work":
		case "sender-street":
		case "sender-city":
		case "sender-postal-code":
		case "sender-country":
		case "sender-state-or-province":
		case "author-name":
		case "author-initials":
		case "chapter":
		case "file-name":
		case "template-name":
		case "sheet-name":
			break;
		case "event-listener":
			break;
		case "initial-creator":
		case "creation-date":
		case "print-date":
		case "generator":
		case "document-statistic":
		case "user-defined":
		case "editing-duration":
		case "editing-cycles":
			break;
		case "config-item":
			break;
		case "page-number":
			break;
		case "page-count":
			break;
		case "time":
			break;
		case "cell-range-source":
			break;
		case "detective":
			break;
		case "operation":
			break;
		case "highlighted-range":
			break;
		case "data-pilot-table":
		case "source-cell-range":
		case "source-service":
		case "data-pilot-field":
		case "data-pilot-level":
		case "data-pilot-subtotals":
		case "data-pilot-subtotal":
		case "data-pilot-members":
		case "data-pilot-member":
		case "data-pilot-display-info":
		case "data-pilot-sort-info":
		case "data-pilot-layout-info":
		case "data-pilot-field-reference":
		case "data-pilot-groups":
		case "data-pilot-group":
		case "data-pilot-group-member":
			break;
		case "rect":
			break;
		case "dde-connection-decls":
		case "dde-connection-decl":
		case "dde-link":
		case "dde-source":
			break;
		case "properties":
			break;
		case "property":
			break;
		case "a":
			if (v[1] !== "/") {
				W = rt(v[0], false);
				if (!W.href) break;
				W.Target = it(W.href);
				delete W.href;
				if (W.Target.charAt(0) == "#" && W.Target.indexOf(".") > -1) {
					j = $s(W.Target.slice(1));
					W.Target = "#" + j[0] + "!" + j[1]
				} else if (W.Target.match(/^\.\.[\\\/]/)) W.Target = W.Target.slice(3)
			}
			break;
		case "table-protection":
			break;
		case "data-pilot-grand-total":
			break;
		case "office-document-common-attrs":
			break;
		default:
			switch (v[2]) {
			case "dc:":
			case "calcext:":
			case "loext:":
			case "ooo:":
			case "chartooo:":
			case "draw:":
			case "style:":
			case "chart:":
			case "form:":
			case "uof:":
			case "表:":
			case "字:":
				break;
			default:
				if (a.WTF) throw new Error(v);
			}
		}
		var ie = {
			Sheets: d,
			SheetNames: p,
			Workbook: $,
		};
		if (a.bookSheets) delete ie.Sheets;
		return ie
	}
	function xo(e, r) {
		r = r || {};
		if (Br(e, "META-INF/manifest.xml")) In(zr(e, "META-INF/manifest.xml"), r);
		var t = $r(e, "styles.xml");
		var a = t && ko(kt(t), r);
		var n = $r(e, "content.xml");
		if (!n) throw new Error("Missing content.xml in ODS / UOF file");
		var i = yo(kt(n), r, a);
		if (Br(e, "meta.xml")) i.Props = Wn(zr(e, "meta.xml"));
		i.bookType = "ods";
		return i
	}
	function _o(e, r) {
		var t = "number",
		a = "",
		n = {
			"style:name": r,
		},
		i = "",
		s = 0;
		e = e.replace(/"[$]"/g, "$");
		e: {
			if (e.indexOf(";") > -1) {
				console.error("Unsupported ODS Style Map exported.  Using first branch of " + e);
				e = e.slice(0, e.indexOf(";"))
			}
			if (e == "@") {
				t = "text";
				a = "<number:text-content/>";
				break e
			}
			if (e.indexOf(/\$/) > -1) {
				t = "currency"
			}
			if (e[s] == '"') {
				i = "";
				while (e[++s] != '"' || e[++s] == '"') i += e[s]; --s;
				if (e[s + 1] == "*") {
					s++;
					a += "<number:fill-character>" + ot(i.replace(/""/g, '"')) + "</number:fill-character>"
				} else {
					a += "<number:text>" + ot(i.replace(/""/g, '"')) + "</number:text>"
				}
				e = e.slice(s + 1);
				s = 0
			}
			var l = e.match(/# (\?+)\/(\?+)/);
			if (l) {
				a += Ot("number:fraction", null, {
					"number:min-integer-digits": 0,
					"number:min-numerator-digits": l[1].length,
					"number:max-denominator-value": Math.max( + l[1].replace(/./g, "9"), +l[2].replace(/./g, "9")),
				});
				break e
			}
			if ((l = e.match(/# (\?+)\/(\d+)/))) {
				a += Ot("number:fraction", null, {
					"number:min-integer-digits": 0,
					"number:min-numerator-digits": l[1].length,
					"number:denominator-value": +l[2],
				});
				break e
			}
			if ((l = e.match(/(\d+)(|\.\d+)%/))) {
				t = "percentage";
				a += Ot("number:number", null, {
					"number:decimal-places": (l[2] && l.length - 1) || 0,
					"number:min-decimal-places": (l[2] && l.length - 1) || 0,
					"number:min-integer-digits": l[1].length,
				}) + "<number:text>%</number:text>";
				break e
			}
			var o = false;
			if (["y", "m", "d"].indexOf(e[0]) > -1) {
				t = "date";
				r: for (; s < e.length; ++s) switch ((i = e[s].toLowerCase())) {
				case "h":
				case "s":
					o = true; --s;
					break r;
				case "m":
					t:
					for (var c = s + 1; c < e.length; ++c) switch (e[c]) {
					case "y":
					case "d":
						break t;
					case "h":
					case "s":
						o = true; --s;
						break r
					}
				case "y":
				case "d":
					while ((e[++s] || "").toLowerCase() == i[0]) i += i[0]; --s;
					switch (i) {
					case "y":
					case "yy":
						a += "<number:year/>";
						break;
					case "yyy":
					case "yyyy":
						a += '<number:year number:style="long"/>';
						break;
					case "mmmmm":
						console.error("ODS has no equivalent of format |mmmmm|");
					case "m":
					case "mm":
					case "mmm":
					case "mmmm":
						a += '<number:month number:style="' + (i.length % 2 ? "short": "long") + '" number:textual="' + (i.length >= 3 ? "true": "false") + '"/>';
						break;
					case "d":
					case "dd":
						a += '<number:day number:style="' + (i.length % 2 ? "short": "long") + '"/>';
						break;
					case "ddd":
					case "dddd":
						a += '<number:day-of-week number:style="' + (i.length % 2 ? "short": "long") + '"/>';
						break
					}
					break;
				case '"':
					while (e[++s] != '"' || e[++s] == '"') i += e[s]; --s;
					a += "<number:text>" + ot(i.slice(1).replace(/""/g, '"')) + "</number:text>";
					break;
				case "\\":
					i = e[++s];
					a += "<number:text>" + ot(i) + "</number:text>";
					break;
				case "/":
				case ":":
					a += "<number:text>" + ot(i) + "</number:text>";
					break;
				default:
					console.error("unrecognized character " + i + " in ODF format " + e)
				}
				if (!o) break e;
				e = e.slice(s + 1);
				s = 0
			}
			if (e.match(/^\[?[hms]/)) {
				if (t == "number") t = "time";
				if (e.match(/\[/)) {
					e = e.replace(/[\[\]]/g, "");
					n["number:truncate-on-overflow"] = "false"
				}
				for (; s < e.length; ++s) switch ((i = e[s].toLowerCase())) {
				case "h":
				case "m":
				case "s":
					while ((e[++s] || "").toLowerCase() == i[0]) i += i[0]; --s;
					switch (i) {
					case "h":
					case "hh":
						a += '<number:hours number:style="' + (i.length % 2 ? "short": "long") + '"/>';
						break;
					case "m":
					case "mm":
						a += '<number:minutes number:style="' + (i.length % 2 ? "short": "long") + '"/>';
						break;
					case "s":
					case "ss":
						if (e[s + 1] == ".") do {
							i += e[s + 1]; ++s
						} while ( e [ s + 1 ] == "0");
						a += '<number:seconds number:style="' + (i.match("ss") ? "long": "short") + '"' + (i.match(/\./) ? ' number:decimal-places="' + (i.match(/0+/) || [""])[0].length + '"': "") + "/>";
						break
					}
					break;
				case '"':
					while (e[++s] != '"' || e[++s] == '"') i += e[s]; --s;
					a += "<number:text>" + ot(i.slice(1).replace(/""/g, '"')) + "</number:text>";
					break;
				case "/":
				case ":":
					a += "<number:text>" + ot(i) + "</number:text>";
					break;
				case "a":
					if (e.slice(s, s + 3).toLowerCase() == "a/p") {
						a += "<number:am-pm/>";
						s += 2;
						break
					}
					if (e.slice(s, s + 5).toLowerCase() == "am/pm") {
						a += "<number:am-pm/>";
						s += 4;
						break
					}
				default:
					console.error("unrecognized character " + i + " in ODF format " + e)
				}
				break e
			}
			if (e.indexOf(/\$/) > -1) {
				t = "currency"
			}
			if (e[0] == "$") {
				a += '<number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>';
				e = e.slice(1);
				s = 0
			}
			s = 0;
			if (e[s] == '"') {
				while (e[++s] != '"' || e[++s] == '"') i += e[s]; --s;
				if (e[s + 1] == "*") {
					s++;
					a += "<number:fill-character>" + ot(i.replace(/""/g, '"')) + "</number:fill-character>"
				} else {
					a += "<number:text>" + ot(i.replace(/""/g, '"')) + "</number:text>"
				}
				e = e.slice(s + 1);
				s = 0
			}
			var f = e.match(/([#0][0#,]*)(\.[0#]*|)(E[+]?0*|)/i);
			if (!f || !f[0]) console.error("Could not find numeric part of " + e);
			else {
				var u = f[1].replace(/,/g, "");
				a += "<number:" + (f[3] ? "scientific-": "") + "number" + ' number:min-integer-digits="' + (u.indexOf("0") == -1 ? "0": u.length - u.indexOf("0")) + '"' + (f[0].indexOf(",") > -1 ? ' number:grouping="true"': "") + ((f[2] && ' number:decimal-places="' + (f[2].length - 1) + '"') || ' number:decimal-places="0"') + (f[3] && f[3].indexOf("+") > -1 ? ' number:forced-exponent-sign="true"': "") + (f[3] ? ' number:min-exponent-digits="' + f[3].match(/0+/)[0].length + '"': "") + "></number:" + (f[3] ? "scientific-": "") + "number>";
				s = f.index + f[0].length
			}
			if (e[s] == '"') {
				i = "";
				while (e[++s] != '"' || e[++s] == '"') i += e[s]; --s;
				a += "<number:text>" + ot(i.replace(/""/g, '"')) + "</number:text>"
			}
		}
		if (!a) {
			console.error("Could not generate ODS number format for |" + e + "|");
			return ""
		}
		return Ot("number:" + t + "-style", a, n)
	}
	function Fo(e) {
		return function r(t) {
			for (var a = 0; a != e.length; ++a) {
				var n = e[a];
				if (t[n[0]] === undefined) t[n[0]] = n[1];
				if (n[2] === "n") t[n[0]] = Number(t[n[0]])
			}
		}
	}
	function Do(e) {
		Fo([["cellNF", false], ["cellHTML", true], ["cellFormula", true], ["cellStyles", false], ["cellText", true], ["cellDates", false], ["sheetStubs", false], ["sheetRows", 0, "n"], ["bookDeps", false], ["bookSheets", false], ["bookProps", false], ["bookFiles", false], ["bookVBA", false], ["password", ""], ["WTF", false], ])(e)
	}
	function Mo(e) {
		if (En.WS.indexOf(e) > -1) return "sheet";
		if (En.CS && e == En.CS) return "chart";
		if (En.DS && e == En.DS) return "dialog";
		if (En.MS && e == En.MS) return "macro";
		return e && e.length ? e: "sheet"
	}
	function No(e, r) {
		if (!e) return 0;
		try {
			e = r.map(function a(r) {
				if (!r.id) r.id = r.strRelID;
				return [r.name, e["!id"][r.id].Target, Mo(e["!id"][r.id].Type)]
			})
		} catch(t) {
			return null
		}
		return ! e || e.length === 0 ? null: e
	}
	function Io(e, r, t, a, n, i, s, l) {
		if (!e || !e["!legdrawel"]) return;
		var o = Gr(e["!legdrawel"].Target, a);
		var c = $r(t, o, true);
		if (c) ds(kt(c), e, l || [])
	}
	function Po(e, r, t, a, n, i, s, l, o, c, f, u) {
		try {
			i[a] = Dn($r(e, t, true), r);
			var h = zr(e, r);
			var d;
			switch (l) {
			case "sheet":
				d = Zl(h, r, n, o, i[a], c, f, u);
				break;
			case "chart":
				d = Kl(h, r, n, o, i[a], c, f, u);
				if (!d || !d["!drawel"]) break;
				var p = Gr(d["!drawel"].Target, r);
				var m = Fn(p);
				var v = us($r(e, p, true), Dn($r(e, m, true), p));
				var g = Gr(v, p);
				var b = Fn(g);
				d = Ml($r(e, g, true), g, o, Dn($r(e, b, true), g), c, d);
				break;
			case "macro":
				d = ql(h, r, n, o, i[a], c, f, u);
				break;
			case "dialog":
				d = Ql(h, r, n, o, i[a], c, f, u);
				break;
			default:
				throw new Error("Unrecognized sheet type " + l);
			}
			s[a] = d;
			var w = [],
			k = [];
			if (i && i[a]) ir(i[a]).forEach(function(t) {
				var n = "";
				if (i[a][t].Type == En.CMNT) {
					n = Gr(i[a][t].Target, r);
					w = to(zr(e, n, true), n, o);
					if (!w || !w.length) return;
					vs(d, w, false)
				}
				if (i[a][t].Type == En.TCMNT) {
					n = Gr(i[a][t].Target, r);
					k = k.concat(ws(zr(e, n, true), o))
				}
			});
			if (k && k.length) vs(d, k, true, o.people || []);
			Io(d, l, e, r, n, o, c, w)
		} catch(y) {
			if (o.WTF) throw y;
		}
	}
	function Ro(e) {
		return e.charAt(0) == "/" ? e.slice(1) : e
	}
	function Lo(e, r) {
		He();
		r = r || {};
		Do(r);
		if (Br(e, "META-INF/manifest.xml")) return xo(e, r);
		if (Br(e, "objectdata.xml")) return xo(e, r);
		if (Br(e, "Index/Document.iwa")) {
			if (typeof Uint8Array == "undefined") throw new Error("NUMBERS file parsing requires Uint8Array support");
			if (typeof parse_numbers_iwa != "undefined") {
				if (e.FileIndex) return parse_numbers_iwa(e, r);
				var t = Qe.utils.cfb_new();
				jr(e).forEach(function(r) {
					Hr(t, r, Wr(e, r))
				});
				return parse_numbers_iwa(t, r)
			}
			throw new Error("Unsupported NUMBERS file");
		}
		if (!Br(e, "[Content_Types].xml")) {
			if (Br(e, "index.xml.gz")) throw new Error("Unsupported NUMBERS 08 file");
			if (Br(e, "index.xml")) throw new Error("Unsupported NUMBERS 09 file");
			var a = Qe.find(e, "Index.zip");
			if (a) {
				r = yr(r);
				delete r.type;
				if (typeof a.content == "string") r.type = "binary";
				if (typeof Bun !== "undefined" && Buffer.isBuffer(a.content)) return Jo(new Uint8Array(a.content), r);
				return Jo(a.content, r)
			}
			throw new Error("Unsupported ZIP file");
		}
		var n = jr(e);
		var i = An($r(e, "[Content_Types].xml"));
		var s = false;
		var l, o;
		if (i.workbooks.length === 0) {
			o = "xl/workbook.xml";
			if (zr(e, o, true)) i.workbooks.push(o)
		}
		if (i.workbooks.length === 0) {
			o = "xl/workbook.bin";
			if (!zr(e, o, true)) throw new Error("Could not find workbook");
			i.workbooks.push(o);
			s = true
		}
		if (i.workbooks[0].slice(-3) == "bin") s = true;
		var c = {};
		var f = {};
		if (!r.bookSheets && !r.bookProps) {
			js = [];
			if (i.sst) try {
				js = ro(zr(e, Ro(i.sst)), i.sst, r)
			} catch(u) {
				if (r.WTF) throw u;
			}
			if (r.cellStyles && i.themes.length) c = is($r(e, i.themes[0].replace(/^\//, ""), true) || "", r);
			if (i.style) f = eo(zr(e, Ro(i.style)), i.style, c, r)
		}
		i.links.map(function(t) {
			try {
				var a = Dn($r(e, Fn(Ro(t))), t);
				return no(zr(e, Ro(t)), a, t, r)
			} catch(n) {}
		});
		var h = Jl(zr(e, Ro(i.workbooks[0])), i.workbooks[0], r);
		var d = {},
		p = "";
		if (i.coreprops.length) {
			p = zr(e, Ro(i.coreprops[0]), true);
			if (p) d = Wn(p);
			if (i.extprops.length !== 0) {
				p = zr(e, Ro(i.extprops[0]), true);
				if (p) Yn(p, d, r)
			}
		}
		var m = {};
		if (!r.bookSheets || r.bookProps) {
			if (i.custprops.length !== 0) {
				p = $r(e, Ro(i.custprops[0]), true);
				if (p) m = Kn(p, r)
			}
		}
		var v = {};
		if (r.bookSheets || r.bookProps) {
			if (h.Sheets) l = h.Sheets.map(function M(e) {
				return e.name
			});
			else if (d.Worksheets && d.SheetNames.length > 0) l = d.SheetNames;
			if (r.bookProps) {
				v.Props = d;
				v.Custprops = m
			}
			if (r.bookSheets && typeof l !== "undefined") v.SheetNames = l;
			if (r.bookSheets ? v.SheetNames: r.bookProps) return v
		}
		l = {};
		var g = {};
		if (r.bookDeps && i.calcchain) g = ao(zr(e, Ro(i.calcchain)), i.calcchain, r);
		var b = 0;
		var w = {};
		var k, y; {
			var x = h.Sheets;
			d.Worksheets = x.length;
			d.SheetNames = [];
			for (var S = 0; S != x.length; ++S) {
				d.SheetNames[S] = x[S].name
			}
		}
		var C = s ? "bin": "xml";
		var _ = i.workbooks[0].lastIndexOf("/");
		var A = (i.workbooks[0].slice(0, _ + 1) + "_rels/" + i.workbooks[0].slice(_ + 1) + ".rels").replace(/^\//, "");
		if (!Br(e, A)) A = "xl/_rels/workbook." + C + ".rels";
		var T = Dn($r(e, A, true), A.replace(/_rels.*/, "s5s"));
		if ((i.metadata || []).length >= 1) {
			r.xlmeta = io(zr(e, Ro(i.metadata[0])), i.metadata[0], r)
		}
		if ((i.people || []).length >= 1) {
			r.people = ys(zr(e, Ro(i.people[0])), r)
		}
		if (T) T = No(T, h.Sheets);
		var E = zr(e, "xl/worksheets/sheet.xml", true) ? 1 : 0;
		e: for (b = 0; b != d.Worksheets; ++b) {
			var F = "sheet";
			if (T && T[b]) {
				k = "xl/" + T[b][1].replace(/[\/]?xl\//, "");
				if (!Br(e, k)) k = T[b][1];
				if (!Br(e, k)) k = A.replace(/_rels\/.*$/, "") + T[b][1];
				F = T[b][2]
			} else {
				k = "xl/worksheets/sheet" + (b + 1 - E) + "." + C;
				k = k.replace(/sheet0\./, "sheet.")
			}
			y = k.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
			if (r && r.sheets != null) switch (typeof r.sheets) {
			case "number":
				if (b != r.sheets) continue e;
				break;
			case "string":
				if (d.SheetNames[b].toLowerCase() != r.sheets.toLowerCase()) continue e;
				break;
			default:
				if (Array.isArray && Array.isArray(r.sheets)) {
					var D = false;
					for (var O = 0; O != r.sheets.length; ++O) {
						if (typeof r.sheets[O] == "number" && r.sheets[O] == b) D = 1;
						if (typeof r.sheets[O] == "string" && r.sheets[O].toLowerCase() == d.SheetNames[b].toLowerCase()) D = 1
					}
					if (!D) continue e
				}
			}
			Po(e, k, y, d.SheetNames[b], b, w, l, F, r, h, c, f)
		}
		v = {
			Directory: i,
			Workbook: h,
			Props: d,
			Custprops: m,
			Deps: g,
			Sheets: l,
			SheetNames: d.SheetNames,
			Strings: js,
			Styles: f,
			Themes: c,
			SSF: yr(q),
		};
		if (r && r.bookFiles) {
			if (e.files) {
				v.keys = n;
				v.files = e.files
			} else {
				v.keys = [];
				v.files = {};
				e.FullPaths.forEach(function(r, t) {
					r = r.replace(/^Root Entry[\/]/, "");
					v.keys.push(r);
					v.files[r] = e.FileIndex[t]
				})
			}
		}
		if (r && r.bookVBA) {
			if (i.vba.length > 0) v.vbaraw = zr(e, Ro(i.vba[0]), true);
			else if (i.defaults && i.defaults.bin === Ss) v.vbaraw = zr(e, "xl/vbaProject.bin", true)
		}
		v.bookType = s ? "xlsb": "xlsx";
		return v
	}
	function Bo(e, r) {
		var t = r || {};
		var a = "Workbook",
		n = Qe.find(e, a);
		try {
			a = "/!DataSpaces/Version";
			n = Qe.find(e, a);
			if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
			parse_DataSpaceVersionInfo(n.content);
			a = "/!DataSpaces/DataSpaceMap";
			n = Qe.find(e, a);
			if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
			var i = parse_DataSpaceMap(n.content);
			if (i.length !== 1 || i[0].comps.length !== 1 || i[0].comps[0].t !== 0 || i[0].name !== "StrongEncryptionDataSpace" || i[0].comps[0].v !== "EncryptedPackage") throw new Error("ECMA-376 Encrypted file bad " + a);
			a = "/!DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace";
			n = Qe.find(e, a);
			if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
			var s = parse_DataSpaceDefinition(n.content);
			if (s.length != 1 || s[0] != "StrongEncryptionTransform") throw new Error("ECMA-376 Encrypted file bad " + a);
			a = "/!DataSpaces/TransformInfo/StrongEncryptionTransform/!Primary";
			n = Qe.find(e, a);
			if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
			parse_Primary(n.content)
		} catch(l) {}
		a = "/EncryptionInfo";
		n = Qe.find(e, a);
		if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
		var o = parse_EncryptionInfo(n.content);
		a = "/EncryptedPackage";
		n = Qe.find(e, a);
		if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
		if (o[0] == 4 && typeof decrypt_agile !== "undefined") return decrypt_agile(o[1], n.content, t.password || "", t);
		if (o[0] == 2 && typeof decrypt_std76 !== "undefined") return decrypt_std76(o[1], n.content, t.password || "", t);
		throw new Error("File is password-protected");
	}
	function $o(e, r) {
		var t = "";
		switch ((r || {}).type || "base64") {
		case "buffer":
			return [e[0], e[1], e[2], e[3], e[4], e[5], e[6], e[7]];
		case "base64":
			t = C(e.slice(0, 12));
			break;
		case "binary":
			t = e;
			break;
		case "array":
			return [e[0], e[1], e[2], e[3], e[4], e[5], e[6], e[7]];
		default:
			throw new Error("Unrecognized type " + ((r && r.type) || "undefined"));
		}
		return [t.charCodeAt(0), t.charCodeAt(1), t.charCodeAt(2), t.charCodeAt(3), t.charCodeAt(4), t.charCodeAt(5), t.charCodeAt(6), t.charCodeAt(7), ]
	}
	function Wo(e, r) {
		if (Qe.find(e, "EncryptedPackage")) return Bo(e, r);
		return parse_xlscfb(e, r)
	}
	function jo(e, r) {
		var t, a = e;
		var n = r || {};
		if (!n.type) n.type = _ && Buffer.isBuffer(e) ? "buffer": "base64";
		t = Xr(a, n);
		return Lo(t, n)
	}
	function Ho(e, r) {
		var t = 0;
		e: while (t < e.length) switch (e.charCodeAt(t)) {
		case 10:
		case 13:
		case 32:
			++t;
			break;
		case 60:
			return parse_xlml(e.slice(t), r);
		default:
			break e
		}
		return ni.to_workbook(e, r)
	}
	function Vo(e, r) {
		var t = "",
		a = $o(e, r);
		switch (r.type) {
		case "base64":
			t = C(e);
			break;
		case "binary":
			t = e;
			break;
		case "buffer":
			t = e.toString("binary");
			break;
		case "array":
			t = kr(e);
			break;
		default:
			throw new Error("Unrecognized type " + r.type);
		}
		if (a[0] == 239 && a[1] == 187 && a[2] == 191) t = kt(t);
		r.type = "binary";
		return Ho(t, r)
	}
	function Xo(e, r) {
		var t = e;
		if (r.type == "base64") t = C(t);
		if (typeof ArrayBuffer !== "undefined" && e instanceof ArrayBuffer) t = new Uint8Array(e);
		t = typeof a !== "undefined" ? a.utils.decode(1200, t.slice(2), "str") : _ && Buffer.isBuffer(e) ? e.slice(2).toString("utf16le") : typeof Uint8Array !== "undefined" && t instanceof Uint8Array ? typeof TextDecoder !== "undefined" ? new TextDecoder("utf-16le").decode(t.slice(2)) : h(t.slice(2)) : u(t.slice(2));
		r.type = "binary";
		return Ho(t, r)
	}
	function Go(e) {
		return ! e.match(/[^\x00-\x7F]/) ? e: yt(e)
	}
	function Yo(e, r, t, a) {
		if (a) {
			t.type = "string";
			return ni.to_workbook(e, t)
		}
		return ni.to_workbook(r, t)
	}
	function Jo(e, r) {
		c();
		var t = r || {};
		if (t.codepage && typeof a === "undefined") console.error("Codepage tables are not loaded.  Non-ASCII characters may not give expected results");
		if (typeof ArrayBuffer !== "undefined" && e instanceof ArrayBuffer) return Jo(new Uint8Array(e), ((t = yr(t)), (t.type = "array"), t));
		if (typeof Uint8Array !== "undefined" && e instanceof Uint8Array && !t.type) t.type = typeof Deno !== "undefined" ? "buffer": "array";
		var n = e,
		i = [0, 0, 0, 0],
		s = false;
		if (t.cellStyles) {
			t.cellNF = true;
			t.sheetStubs = true
		}
		Hs = {};
		if (t.dateNF) Hs.dateNF = t.dateNF;
		if (!t.type) t.type = _ && Buffer.isBuffer(e) ? "buffer": "base64";
		if (t.type == "file") {
			t.type = _ ? "buffer": "binary";
			n = nr(e);
			if (typeof Uint8Array !== "undefined" && !_) t.type = "array"
		}
		if (t.type == "string") {
			s = true;
			t.type = "binary";
			t.codepage = 65001;
			n = Go(e)
		}
		if (t.type == "array" && typeof Uint8Array !== "undefined" && e instanceof Uint8Array && typeof ArrayBuffer !== "undefined") {
			var l = new ArrayBuffer(3),
			o = new Uint8Array(l);
			o.foo = "bar";
			if (!o.foo) {
				t = yr(t);
				t.type = "array";
				return Jo(I(n), t)
			}
		}
		switch ((i = $o(n, t))[0]) {
		case 208:
			if (i[1] === 207 && i[2] === 17 && i[3] === 224 && i[4] === 161 && i[5] === 177 && i[6] === 26 && i[7] === 225) return Wo(Qe.read(n, t), t);
			break;
		case 9:
			if (i[1] <= 8) return parse_xlscfb(n, t);
			break;
		case 60:
			return parse_xlml(n, t);
		case 73:
			if (i[1] === 73 && i[2] === 42 && i[3] === 0) throw new Error("TIFF Image File is not a spreadsheet");
			if (i[1] === 68) return ii(n, t);
			break;
		case 84:
			if (i[1] === 65 && i[2] === 66 && i[3] === 76) return ti.to_workbook(n, t);
			break;
		case 80:
			return i[1] === 75 && i[2] < 9 && i[3] < 9 ? jo(n, t) : Yo(e, n, t, s);
		case 239:
			return i[3] === 60 ? parse_xlml(n, t) : Yo(e, n, t, s);
		case 255:
			if (i[1] === 254) {
				return Xo(n, t)
			} else if (i[1] === 0 && i[2] === 2 && i[3] === 0) return WK_.to_workbook(n, t);
			break;
		case 0:
			if (i[1] === 0) {
				if (i[2] >= 2 && i[3] === 0) return WK_.to_workbook(n, t);
				if (i[2] === 0 && (i[3] === 8 || i[3] === 9)) return WK_.to_workbook(n, t)
			}
			break;
		case 3:
		case 131:
		case 139:
		case 140:
			return ei.to_workbook(n, t);
		case 123:
			if (i[1] === 92 && i[2] === 114 && i[3] === 116) return rtf_to_workbook(n, t);
			break;
		case 10:
		case 13:
		case 32:
			return Vo(n, t);
		case 137:
			if (i[1] === 80 && i[2] === 78 && i[3] === 71) throw new Error("PNG Image File is not a spreadsheet");
			break;
		case 8:
			if (i[1] === 231) throw new Error("Unsupported Multiplan 1.x file!");
			break;
		case 12:
			if (i[1] === 236) throw new Error("Unsupported Multiplan 2.x file!");
			if (i[1] === 237) throw new Error("Unsupported Multiplan 3.x file!");
			break
		}
		if (Qn.indexOf(i[0]) > -1 && i[2] <= 12 && i[3] <= 31) return ei.to_workbook(n, t);
		return Yo(e, n, t, s)
	}
	function hc(e, r, t, a, n, i, s) {
		var l = Ma(t);
		var o = s.defval,
		c = s.raw || !Object.prototype.hasOwnProperty.call(s, "raw");
		var f = true,
		u = e["!data"] != null;
		var h = n === 1 ? [] : {};
		if (n !== 1) {
			if (Object.defineProperty) try {
				Object.defineProperty(h, "__rowNum__", {
					value: t,
					enumerable: false,
				})
			} catch(d) {
				h.__rowNum__ = t
			} else h.__rowNum__ = t
		}
		if (!u || e["!data"][t]) for (var p = r.s.c; p <= r.e.c; ++p) {
			var m = u ? (e["!data"][t] || [])[p] : e[a[p] + l];
			if (m == null || m.t === undefined) {
				if (o === undefined) continue;
				if (i[p] != null) {
					h[i[p]] = o
				}
				continue
			}
			var v = m.v;
			switch (m.t) {
			case "z":
				if (v == null) break;
				continue;
			case "e":
				v = v == 0 ? null: void 0;
				break;
			case "s":
			case "b":
			case "n":
				if (!m.z || !Re(m.z)) break;
				v = pr(v);
				if (typeof v == "number") break;
			case "d":
				if (! (s && (s.UTC || s.raw === false))) v = Nr(new Date(v));
				break;
			default:
				throw new Error("unrecognized type " + m.t);
			}
			if (i[p] != null) {
				if (v == null) {
					if (m.t == "e" && v === null) h[i[p]] = null;
					else if (o !== undefined) h[i[p]] = o;
					else if (c && v === null) h[i[p]] = null;
					else continue
				} else {
					h[i[p]] = (m.t === "n" && typeof s.rawNumbers === "boolean" ? s.rawNumbers: c) ? v: Ya(m, v, s)
				}
				if (v != null) f = false
			}
		}
		return {
			row: h,
			isempty: f,
		}
	}
	function dc(e, r) {
		if (e == null || e["!ref"] == null) return [];
		var t = {
			t: "n",
			v: 0,
		},
		a = 0,
		n = 1,
		i = [],
		s = 0,
		l = "";
		var o = {
			s: {
				r: 0,
				c: 0,
			},
			e: {
				r: 0,
				c: 0,
			},
		};
		var c = r || {};
		var f = c.range != null ? c.range: e["!ref"];
		if (c.header === 1) a = 1;
		else if (c.header === "A") a = 2;
		else if (Array.isArray(c.header)) a = 3;
		else if (c.header == null) a = 0;
		switch (typeof f) {
		case "string":
			o = Xa(f);
			break;
		case "number":
			o = Xa(e["!ref"]);
			o.s.r = f;
			break;
		default:
			o = f
		}
		if (a > 0) n = 0;
		var u = Ma(o.s.r);
		var h = [];
		var d = [];
		var p = 0,
		m = 0;
		var v = e["!data"] != null;
		var g = o.s.r,
		b = 0;
		var w = {};
		if (v && !e["!data"][g]) e["!data"][g] = [];
		var k = (c.skipHidden && e["!cols"]) || [];
		var y = (c.skipHidden && e["!rows"]) || [];
		for (b = o.s.c; b <= o.e.c; ++b) {
			if ((k[b] || {}).hidden) continue;
			h[b] = Ra(b);
			t = v ? e["!data"][g][b] : e[h[b] + u];
			switch (a) {
			case 1:
				i[b] = b - o.s.c;
				break;
			case 2:
				i[b] = h[b];
				break;
			case 3:
				i[b] = c.header[b - o.s.c];
				break;
			default:
				if (t == null) t = {
					w: "__EMPTY",
					t: "s",
				};
				l = s = Ya(t, null, c);
				m = w[s] || 0;
				if (!m) w[s] = 1;
				else {
					do {
						l = s + "_" + m++
					} while ( w [ l ]);
					w[s] = m;
					w[l] = 1
				}
				i[b] = l
			}
		}
		for (g = o.s.r + n; g <= o.e.r; ++g) {
			if ((y[g] || {}).hidden) continue;
			var x = hc(e, o, g, h, a, i, c);
			if (x.isempty === false || (a === 1 ? c.blankrows !== false: !!c.blankrows)) d[p++] = x.row
		}
		d.length = p;
		return d
	}
	var pc = /"/g;
	function mc(e, r, t, a, n, i, s, l) {
		var o = true;
		var c = [],
		f = "",
		u = Ma(t);
		var h = e["!data"] != null;
		var d = (h && e["!data"][t]) || [];
		for (var p = r.s.c; p <= r.e.c; ++p) {
			if (!a[p]) continue;
			var m = h ? d[p] : e[a[p] + u];
			if (m == null) f = "";
			else if (m.v != null) {
				o = false;
				f = "" + (l.rawNumbers && m.t == "n" ? m.v: Ya(m, null, l));
				for (var v = 0,
				g = 0; v !== f.length; ++v) if ((g = f.charCodeAt(v)) === n || g === i || g === 34 || l.forceQuotes) {
					f = '"' + f.replace(pc, '""') + '"';
					break
				}
				if (f == "ID") f = '"ID"'
			} else if (m.f != null && !m.F) {
				o = false;
				f = "=" + m.f;
				if (f.indexOf(",") >= 0) f = '"' + f.replace(pc, '""') + '"'
			} else f = "";
			c.push(f)
		}
		if (l.blankrows === false && o) return null;
		return c.join(s)
	}
	function vc(e, r) {
		var t = [];
		var a = r == null ? {}: r;
		if (e == null || e["!ref"] == null) return "";
		var n = Xa(e["!ref"]);
		var i = a.FS !== undefined ? a.FS: ",",
		s = i.charCodeAt(0);
		var l = a.RS !== undefined ? a.RS: "\n",
		o = l.charCodeAt(0);
		var c = new RegExp((i == "|" ? "\\|": i) + "+$");
		var f = "",
		u = [];
		var h = (a.skipHidden && e["!cols"]) || [];
		var d = (a.skipHidden && e["!rows"]) || [];
		for (var p = n.s.c; p <= n.e.c; ++p) if (! (h[p] || {}).hidden) u[p] = Ra(p);
		var m = 0;
		for (var v = n.s.r; v <= n.e.r; ++v) {
			if ((d[v] || {}).hidden) continue;
			f = mc(e, n, v, u, s, o, i, a);
			if (f == null) {
				continue
			}
			if (a.strip) f = f.replace(c, "");
			if (f || a.blankrows !== false) t.push((m++?l: "") + f)
		}
		return t.join("")
	}
	function gc(e, r) {
		if (!r) r = {};
		r.FS = "\t";
		r.RS = "\n";
		var t = vc(e, r);
		if (typeof a == "undefined" || r.type == "string") return t;
		var n = a.utils.encode(1200, t, "str");
		return String.fromCharCode(255) + String.fromCharCode(254) + n
	}
	function bc(e) {
		var r = "",
		t, a = "";
		if (e == null || e["!ref"] == null) return [];
		var n = Xa(e["!ref"]),
		i = "",
		s = [],
		l;
		var o = [];
		var c = e["!data"] != null;
		for (l = n.s.c; l <= n.e.c; ++l) s[l] = Ra(l);
		for (var f = n.s.r; f <= n.e.r; ++f) {
			i = Ma(f);
			for (l = n.s.c; l <= n.e.c; ++l) {
				r = s[l] + i;
				t = c ? (e["!data"][f] || [])[l] : e[r];
				a = "";
				if (t === undefined) continue;
				else if (t.F != null) {
					r = t.F;
					if (!t.f) continue;
					a = t.f;
					if (r.indexOf(":") == -1) r = r + ":" + r
				}
				if (t.f != null) a = t.f;
				else if (t.t == "z") continue;
				else if (t.t == "n" && t.v != null) a = "" + t.v;
				else if (t.t == "b") a = t.v ? "TRUE": "FALSE";
				else if (t.w !== undefined) a = "'" + t.w;
				else if (t.v === undefined) continue;
				else if (t.t == "s") a = "'" + t.v;
				else a = "" + t.v;
				o[o.length] = r + "=" + a
			}
		}
		return o
	}
	function wc(e, r, t) {
		var a = t || {};
		var n = e ? e["!data"] != null: a.dense;
		if (b != null && n == null) n = b;
		var i = +!a.skipHeader;
		var s = e || {};
		if (!e && n) s["!data"] = [];
		var l = 0,
		o = 0;
		if (s && a.origin != null) {
			if (typeof a.origin == "number") l = a.origin;
			else {
				var c = typeof a.origin == "string" ? za(a.origin) : a.origin;
				l = c.r;
				o = c.c
			}
		}
		var f = {
			s: {
				c: 0,
				r: 0,
			},
			e: {
				c: o,
				r: l + r.length - 1 + i,
			},
		};
		if (s["!ref"]) {
			var u = Xa(s["!ref"]);
			f.e.c = Math.max(f.e.c, u.e.c);
			f.e.r = Math.max(f.e.r, u.e.r);
			if (l == -1) {
				l = u.e.r + 1;
				f.e.r = l + r.length - 1 + i
			}
		} else {
			if (l == -1) {
				l = 0;
				f.e.r = r.length - 1 + i
			}
		}
		var h = a.header || [],
		d = 0;
		var p = [];
		r.forEach(function(e, r) {
			if (n && !s["!data"][l + r + i]) s["!data"][l + r + i] = [];
			if (n) p = s["!data"][l + r + i];
			ir(e).forEach(function(t) {
				if ((d = h.indexOf(t)) == -1) h[(d = h.length)] = t;
				var c = e[t];
				var f = "z";
				var u = "";
				var m = n ? "": Ra(o + d) + Ma(l + r + i);
				var v = n ? p[o + d] : s[m];
				if (c && typeof c === "object" && !(c instanceof Date)) {
					if (n) p[o + d] = c;
					else s[m] = c
				} else {
					if (typeof c == "number") f = "n";
					else if (typeof c == "boolean") f = "b";
					else if (typeof c == "string") f = "s";
					else if (c instanceof Date) {
						f = "d";
						if (!a.UTC) c = Ir(c);
						if (!a.cellDates) {
							f = "n";
							c = dr(c)
						}
						u = v != null && v.z && Re(v.z) ? v.z: a.dateNF || q[14]
					} else if (c === null && a.nullError) {
						f = "e";
						c = 0
					}
					if (!v) {
						if (!n) s[m] = v = {
							t: f,
							v: c,
						};
						else p[o + d] = v = {
							t: f,
							v: c,
						}
					} else {
						v.t = f;
						v.v = c;
						delete v.w;
						delete v.R;
						if (u) v.z = u
					}
					if (u) v.z = u
				}
			})
		});
		f.e.c = Math.max(f.e.c, o + h.length - 1);
		var m = Ma(l);
		if (n && !s["!data"][l]) s["!data"][l] = [];
		if (i) for (d = 0; d < h.length; ++d) {
			if (n) s["!data"][l][d + o] = {
				t: "s",
				v: h[d],
			};
			else s[Ra(d + o) + m] = {
				t: "s",
				v: h[d],
			}
		}
		s["!ref"] = ja(f);
		return s
	}
	
	function yc(e, r, t) {
		if (typeof r == "string") {
			if (e["!data"] != null) {
				var a = za(r);
				if (!e["!data"][a.r]) e["!data"][a.r] = [];
				return (e["!data"][a.r][a.c] || (e["!data"][a.r][a.c] = {
					t: "z",
				}))
			}
			return (e[r] || (e[r] = {
				t: "z",
			}))
		}
		if (typeof r != "number") return yc(e, $a(r));
		return yc(e, Ra(t || 0) + Ma(r))
	}
	function xc(e, r) {
		if (typeof r == "number") {
			if (r >= 0 && e.SheetNames.length > r) return r;
			throw new Error("Cannot find sheet # " + r);
		} else if (typeof r == "string") {
			var t = e.SheetNames.indexOf(r);
			if (t > -1) return t;
			throw new Error("Cannot find sheet name |" + r + "|");
		} else throw new Error("Cannot find sheet |" + r + "|");
	}
	function Sc(e, r) {
		var t = {
			SheetNames: [],
			Sheets: {},
		};
		if (e) Cc(t, e, r || "Sheet1");
		return t
	}
	function Cc(e, r, t, a) {
		var n = 1;
		if (!t) for (; n <= 65535; ++n, t = undefined) if (e.SheetNames.indexOf((t = "Sheet" + n)) == -1) break;
		if (!t || e.SheetNames.length >= 65535) throw new Error("Too many worksheets");
		if (a && e.SheetNames.indexOf(t) >= 0) {
			var i = t.match(/(^.*?)(\d+)$/);
			n = (i && +i[2]) || 0;
			var s = (i && i[1]) || t;
			for (++n; n <= 65535; ++n) if (e.SheetNames.indexOf((t = s + n)) == -1) break
		}
		jl(t);
		if (e.SheetNames.indexOf(t) >= 0) throw new Error("Worksheet with name |" + t + "| already exists!");
		e.SheetNames.push(t);
		e.Sheets[t] = r;
		return t
	}
	function _c(e, r, t) {
		if (!e.Workbook) e.Workbook = {};
		if (!e.Workbook.Sheets) e.Workbook.Sheets = [];
		var a = xc(e, r);
		if (!e.Workbook.Sheets[a]) e.Workbook.Sheets[a] = {};
		switch (t) {
		case 0:
		case 1:
		case 2:
			break;
		default:
			throw new Error("Bad sheet visibility setting " + t);
		}
		e.Workbook.Sheets[a].Hidden = t
	}
	function Ac(e, r) {
		e.z = r;
		return e
	}
	function Tc(e, r, t) {
		if (!r) {
			delete e.l
		} else {
			e.l = {
				Target: r,
			};
			if (t) e.l.Tooltip = t
		}
		return e
	}
	function Ec(e, r, t) {
		return Tc(e, "#" + r, t)
	}
	function Fc(e, r, t) {
		if (!e.c) e.c = [];
		e.c.push({
			t: r,
			a: t || "SheetJS",
		})
	}
	function Dc(e, r, t, a) {
		var n = typeof r != "string" ? r: Xa(r);
		var i = typeof r == "string" ? r: ja(r);
		for (var s = n.s.r; s <= n.e.r; ++s) for (var l = n.s.c; l <= n.e.c; ++l) {
			var o = yc(e, s, l);
			o.t = "n";
			o.F = i;
			delete o.v;
			if (s == n.s.r && l == n.s.c) {
				o.f = t;
				if (a) o.D = true
			}
		}
		var c = Wa(e["!ref"]);
		if (c.s.r > n.s.r) c.s.r = n.s.r;
		if (c.s.c > n.s.c) c.s.c = n.s.c;
		if (c.e.r < n.e.r) c.e.r = n.e.r;
		if (c.e.c < n.e.c) c.e.c = n.e.c;
		e["!ref"] = ja(c);
		return e
	}
	var Oc = {
		encode_col: Ra,
		encode_row: Ma,
		encode_cell: $a,
		encode_range: ja,
		decode_col: Pa,
		decode_row: Oa,
		split_cell: Ua,
		decode_cell: za,
		decode_range: Wa,
		format_cell: Ya,
		sheet_new: Za,
		sheet_add_aoa: Ka,
		sheet_add_json: wc,
		sheet_add_dom: po,
		table_to_book: vo,
		sheet_to_csv: vc,
		sheet_to_txt: gc,
		sheet_to_json: dc,
		sheet_to_html: ho,
		sheet_to_formulae: bc,
		sheet_to_row_object_array: dc,
		sheet_get_cell: yc,
		book_new: Sc,
		book_append_sheet: Cc,
		book_set_sheet_visibility: _c,
		cell_set_number_format: Ac,
		cell_set_hyperlink: Tc,
		cell_set_internal_link: Ec,
		cell_add_comment: Fc,
		sheet_set_array_formula: Dc,
		consts: {
			SHEET_VISIBLE: 0,
			SHEET_HIDDEN: 1,
			SHEET_VERY_HIDDEN: 2,
		},
	};
	e.read = Jo;
	e.utils = Oc;
	if (typeof require !== "undefined") {
		var Mc = undefined;
		if ((Mc || {}).Readable) set_readable(Mc.Readable);
		try {
			er = undefined
		} catch(Ns) {}
	}
}
if (typeof exports !== "undefined") make_xlsx_lib(exports);
else if (typeof module !== "undefined" && module.exports) make_xlsx_lib(module.exports);
else if (typeof define === "function" && define.amd) define("xlsx",
function() {
	if (!XLSX.version) make_xlsx_lib(XLSX);
	return XLSX
});
else make_xlsx_lib(XLSX);
if (typeof window !== "undefined" && !window.XLSX) try {
	window.XLSX = XLSX
} catch(e) {}
