/*! xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
var XLSX = {};
function make_xlsx_lib(e) {
	e.version = "0.20.1";
	var r = 1200,
	t = 1252;
	var a;
	var n = [874, 932, 936, 949, 950, 1250, 1251, 1252, 1253, 1254, 1255, 1256, 1257, 1258, 1e4];
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
		69 : 6969
	};
	var s = function(e) {
		if (n.indexOf(e) == -1) return;
		t = i[0] = e
	};
	function f() {
		s(1252)
	}
	var l = function(e) {
		r = e;
		s(e)
	};
	function o() {
		l(1200);
		f()
	}
	function c(e) {
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
	var v = function(e) {
		var r = e.charCodeAt(0),
		t = e.charCodeAt(1);
		if (r == 255 && t == 254) return u(e.slice(2));
		if (r == 254 && t == 255) return d(e.slice(2));
		if (r == 65279) return e.slice(1);
		return e
	};
	var p = function oT(e) {
		return String.fromCharCode(e)
	};
	var m = function cT(e) {
		return String.fromCharCode(e)
	};
	function b(e) {
		a = e;
		l = function(e) {
			r = e;
			s(e)
		};
		v = function(e) {
			if (e.charCodeAt(0) === 255 && e.charCodeAt(1) === 254) {
				return a.utils.decode(1200, c(e.slice(2)))
			}
			return e
		};
		p = function n(e) {
			if (r === 1200) return String.fromCharCode(e);
			return a.utils.decode(r, [e & 255, e >> 8])[0]
		};
		m = function i(e) {
			return a.utils.decode(t, [e])[0]
		};
		la()
	}
	var g = null;
	var w = true;
	var k = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
	function T(e) {
		var r = "";
		var t = 0,
		a = 0,
		n = 0,
		i = 0,
		s = 0,
		f = 0,
		l = 0;
		for (var o = 0; o < e.length;) {
			t = e.charCodeAt(o++);
			i = t >> 2;
			a = e.charCodeAt(o++);
			s = (t & 3) << 4 | a >> 4;
			n = e.charCodeAt(o++);
			f = (a & 15) << 2 | n >> 6;
			l = n & 63;
			if (isNaN(a)) {
				f = l = 64
			} else if (isNaN(n)) {
				l = 64
			}
			r += k.charAt(i) + k.charAt(s) + k.charAt(f) + k.charAt(l)
		}
		return r
	}
	function y(e) {
		var r = "";
		var t = 0,
		a = 0,
		n = 0,
		i = 0,
		s = 0,
		f = 0,
		l = 0;
		for (var o = 0; o < e.length;) {
			t = e.charCodeAt(o++);
			if (t > 255) t = 95;
			i = t >> 2;
			a = e.charCodeAt(o++);
			if (a > 255) a = 95;
			s = (t & 3) << 4 | a >> 4;
			n = e.charCodeAt(o++);
			if (n > 255) n = 95;
			f = (a & 15) << 2 | n >> 6;
			l = n & 63;
			if (isNaN(a)) {
				f = l = 64
			} else if (isNaN(n)) {
				l = 64
			}
			r += k.charAt(i) + k.charAt(s) + k.charAt(f) + k.charAt(l)
		}
		return r
	}
	function E(e) {
		var r = "";
		var t = 0,
		a = 0,
		n = 0,
		i = 0,
		s = 0,
		f = 0,
		l = 0;
		for (var o = 0; o < e.length;) {
			t = e[o++];
			i = t >> 2;
			a = e[o++];
			s = (t & 3) << 4 | a >> 4;
			n = e[o++];
			f = (a & 15) << 2 | n >> 6;
			l = n & 63;
			if (isNaN(a)) {
				f = l = 64
			} else if (isNaN(n)) {
				l = 64
			}
			r += k.charAt(i) + k.charAt(s) + k.charAt(f) + k.charAt(l)
		}
		return r
	}
	function _(e) {
		var r = "";
		var t = 0,
		a = 0,
		n = 0,
		i = 0,
		s = 0,
		f = 0,
		l = 0;
		e = e.replace(/^data:([^\/]+\/[^\/]+)?;base64\,/, "").replace(/[^\w\+\/\=]/g, "");
		for (var o = 0; o < e.length;) {
			i = k.indexOf(e.charAt(o++));
			s = k.indexOf(e.charAt(o++));
			t = i << 2 | s >> 4;
			r += String.fromCharCode(t);
			f = k.indexOf(e.charAt(o++));
			a = (s & 15) << 4 | f >> 2;
			if (f !== 64) {
				r += String.fromCharCode(a)
			}
			l = k.indexOf(e.charAt(o++));
			n = (f & 3) << 6 | l;
			if (l !== 64) {
				r += String.fromCharCode(n)
			}
		}
		return r
	}
	var S = function() {
		return typeof Buffer !== "undefined" && typeof undefined !== "undefined" && typeof {} !== "undefined" && !!{}.node
	} ();
	var x = function() {
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
	} ();
	var A = function() {
		if (typeof Buffer === "undefined") return false;
		var e = x([65, 0]);
		if (!e) return false;
		var r = e.toString("utf16le");
		return r.length == 1
	} ();
	function C(e) {
		if (S) return Buffer.alloc ? Buffer.alloc(e) : new Buffer(e);
		return typeof Uint8Array != "undefined" ? new Uint8Array(e) : new Array(e)
	}
	function R(e) {
		if (S) return Buffer.allocUnsafe ? Buffer.allocUnsafe(e) : new Buffer(e);
		return typeof Uint8Array != "undefined" ? new Uint8Array(e) : new Array(e)
	}
	var O = function uT(e) {
		if (S) return x(e, "binary");
		return e.split("").map(function(e) {
			return e.charCodeAt(0) & 255
		})
	};
	function I(e) {
		if (typeof ArrayBuffer === "undefined") return O(e);
		var r = new ArrayBuffer(e.length),
		t = new Uint8Array(r);
		for (var a = 0; a != e.length; ++a) t[a] = e.charCodeAt(a) & 255;
		return r
	}
	function N(e) {
		if (Array.isArray(e)) return e.map(function(e) {
			return String.fromCharCode(e)
		}).join("");
		var r = [];
		for (var t = 0; t < e.length; ++t) r[t] = String.fromCharCode(e[t]);
		return r.join("")
	}
	function F(e) {
		if (typeof Uint8Array === "undefined") throw new Error("Unsupported");
		return new Uint8Array(e)
	}
	function D(e) {
		if (typeof ArrayBuffer == "undefined") throw new Error("Unsupported");
		if (e instanceof ArrayBuffer) return D(new Uint8Array(e));
		var r = new Array(e.length);
		for (var t = 0; t < e.length; ++t) r[t] = e[t];
		return r
	}
	var P = S ?
	function(e) {
		return Buffer.concat(e.map(function(e) {
			return Buffer.isBuffer(e) ? e: x(e)
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
				else if (typeof e[r] == "string") a.set(new Uint8Array(O(e[r])), t);
				else a.set(new Uint8Array(e[r]), t)
			}
			return a
		}
		return [].concat.apply([], e.map(function(e) {
			return Array.isArray(e) ? e: [].slice.call(e)
		}))
	};
	function L(e) {
		var r = [],
		t = 0,
		a = e.length + 250;
		var n = C(e.length + 255);
		for (var i = 0; i < e.length; ++i) {
			var s = e.charCodeAt(i);
			if (s < 128) n[t++] = s;
			else if (s < 2048) {
				n[t++] = 192 | s >> 6 & 31;
				n[t++] = 128 | s & 63
			} else if (s >= 55296 && s < 57344) {
				s = (s & 1023) + 64;
				var f = e.charCodeAt(++i) & 1023;
				n[t++] = 240 | s >> 8 & 7;
				n[t++] = 128 | s >> 2 & 63;
				n[t++] = 128 | f >> 6 & 15 | (s & 3) << 4;
				n[t++] = 128 | f & 63
			} else {
				n[t++] = 224 | s >> 12 & 15;
				n[t++] = 128 | s >> 6 & 63;
				n[t++] = 128 | s & 63
			}
			if (t > a) {
				r.push(n.slice(0, t));
				t = 0;
				n = C(65535);
				a = 65530
			}
		}
		r.push(n.slice(0, t));
		return P(r)
	}
	var M = /\u0000/g,
	U = /[\u0001-\u0006]/g;
	function B(e) {
		var r = "",
		t = e.length - 1;
		while (t >= 0) r += e.charAt(t--);
		return r
	}
	function W(e, r) {
		var t = "" + e;
		return t.length >= r ? t: yr("0", r - t.length) + t
	}
	function z(e, r) {
		var t = "" + e;
		return t.length >= r ? t: yr(" ", r - t.length) + t
	}
	function H(e, r) {
		var t = "" + e;
		return t.length >= r ? t: t + yr(" ", r - t.length)
	}
	function V(e, r) {
		var t = "" + Math.round(e);
		return t.length >= r ? t: yr("0", r - t.length) + t
	}
	function X(e, r) {
		var t = "" + e;
		return t.length >= r ? t: yr("0", r - t.length) + t
	}
	var G = Math.pow(2, 32);
	function j(e, r) {
		if (e > G || e < -G) return V(e, r);
		var t = Math.round(e);
		return X(t, r)
	}
	function K(e, r) {
		r = r || 0;
		return e.length >= 7 + r && (e.charCodeAt(r) | 32) === 103 && (e.charCodeAt(r + 1) | 32) === 101 && (e.charCodeAt(r + 2) | 32) === 110 && (e.charCodeAt(r + 3) | 32) === 101 && (e.charCodeAt(r + 4) | 32) === 114 && (e.charCodeAt(r + 5) | 32) === 97 && (e.charCodeAt(r + 6) | 32) === 108
	}
	var Y = [["Sun", "Sunday"], ["Mon", "Monday"], ["Tue", "Tuesday"], ["Wed", "Wednesday"], ["Thu", "Thursday"], ["Fri", "Friday"], ["Sat", "Saturday"]];
	var Z = [["J", "Jan", "January"], ["F", "Feb", "February"], ["M", "Mar", "March"], ["A", "Apr", "April"], ["M", "May", "May"], ["J", "Jun", "June"], ["J", "Jul", "July"], ["A", "Aug", "August"], ["S", "Sep", "September"], ["O", "Oct", "October"], ["N", "Nov", "November"], ["D", "Dec", "December"]];
	function J(e) {
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
		56 : '"上午/下午 "hh"時"mm"分"ss"秒 "'
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
		82 : 0
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
		44 : '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
	};
	function re(e, r, t) {
		var a = e < 0 ? -1 : 1;
		var n = e * a;
		var i = 0,
		s = 1,
		f = 0;
		var l = 1,
		o = 0,
		c = 0;
		var u = Math.floor(n);
		while (o < r) {
			u = Math.floor(n);
			f = u * s + i;
			c = u * o + l;
			if (n - u < 5e-8) break;
			n = 1 / (n - u);
			i = s;
			s = f;
			l = o;
			o = c
		}
		if (c > r) {
			if (o > r) {
				c = l;
				f = i
			} else {
				c = o;
				f = s
			}
		}
		if (!t) return [0, a * f, c];
		var h = Math.floor(a * f / c);
		return [h, a * f - h * c, c]
	}
	function te(e) {
		var r = e.toPrecision(16);
		if (r.indexOf("e") > -1) {
			var t = r.slice(0, r.indexOf("e"));
			t = t.indexOf(".") > -1 ? t.slice(0, t.slice(0, 2) == "0." ? 17 : 16) : t.slice(0, 15) + yr("0", t.length - 15);
			return t + r.slice(r.indexOf("e"))
		}
		var a = r.indexOf(".") > -1 ? r.slice(0, r.slice(0, 2) == "0." ? 17 : 16) : r.slice(0, 15) + yr("0", r.length - 15);
		return Number(a)
	}
	function ae(e, r, t) {
		if (e > 2958465 || e < 0) return null;
		e = te(e);
		var a = e | 0,
		n = Math.floor(86400 * (e - a)),
		i = 0;
		var s = [];
		var f = {
			D: a,
			T: n,
			u: 86400 * (e - a) - n,
			y: 0,
			m: 0,
			d: 0,
			H: 0,
			M: 0,
			S: 0,
			q: 0
		};
		if (Math.abs(f.u) < 1e-6) f.u = 0;
		if (r && r.date1904) a += 1462;
		if (f.u > .9999) {
			f.u = 0;
			if (++n == 86400) {
				f.T = n = 0; ++a; ++f.D
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
			var l = new Date(1900, 0, 1);
			l.setDate(l.getDate() + a - 1);
			s = [l.getFullYear(), l.getMonth() + 1, l.getDate()];
			i = l.getDay();
			if (a < 60) i = (i + 6) % 7;
			if (t) i = ce(l, s)
		}
		f.y = s[0];
		f.m = s[1];
		f.d = s[2];
		f.S = n % 60;
		n = Math.floor(n / 60);
		f.M = n % 60;
		n = Math.floor(n / 60);
		f.H = n;
		f.q = i;
		return f
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
	function fe(e) {
		var r = ne(e.toFixed(11));
		return r.length > (e < 0 ? 12 : 11) || r === "0" || r === "-0" ? e.toPrecision(6) : r
	}
	function le(e) {
		var r = Math.floor(Math.log(Math.abs(e)) * Math.LOG10E),
		t;
		if (r >= -4 && r <= -1) t = e.toPrecision(10 + r);
		else if (Math.abs(r) <= 9) t = se(e);
		else if (r === 10) t = e.toFixed(10).substr(0, 12);
		else t = fe(e);
		return ne(ie(t.toUpperCase()))
	}
	function oe(e, r) {
		switch (typeof e) {
		case "string":
			return e;
		case "boolean":
			return e ? "TRUE": "FALSE";
		case "number":
			return (e | 0) === e ? e.toString(10) : le(e);
		case "undefined":
			return "";
		case "object":
			if (e == null) return "";
			if (e instanceof Date) return ze(14, dr(e, r && r.date1904), r);
		}
		throw new Error("unsupported value in General format: " + e)
	}
	function ce(e, r) {
		r[0] -= 581;
		var t = e.getDay();
		if (e < 60) t = (t + 6) % 7;
		return t
	}
	function ue(e, r, t, a) {
		var n = "",
		i = 0,
		s = 0,
		f = t.y,
		l, o = 0;
		switch (e) {
		case 98:
			f = t.y + 543;
		case 121:
			switch (r.length) {
			case 1:
				;
			case 2:
				l = f % 100;
				o = 2;
				break;
			default:
				l = f % 1e4;
				o = 4;
				break;
			}
			break;
		case 109:
			switch (r.length) {
			case 1:
				;
			case 2:
				l = t.m;
				o = r.length;
				break;
			case 3:
				return Z[t.m - 1][1];
			case 5:
				return Z[t.m - 1][0];
			default:
				return Z[t.m - 1][2];
			}
			break;
		case 100:
			switch (r.length) {
			case 1:
				;
			case 2:
				l = t.d;
				o = r.length;
				break;
			case 3:
				return Y[t.q][0];
			default:
				return Y[t.q][1];
			}
			break;
		case 104:
			switch (r.length) {
			case 1:
				;
			case 2:
				l = 1 + (t.H + 11) % 12;
				o = r.length;
				break;
			default:
				throw "bad hour format: " + r;
			}
			break;
		case 72:
			switch (r.length) {
			case 1:
				;
			case 2:
				l = t.H;
				o = r.length;
				break;
			default:
				throw "bad hour format: " + r;
			}
			break;
		case 77:
			switch (r.length) {
			case 1:
				;
			case 2:
				l = t.M;
				o = r.length;
				break;
			default:
				throw "bad minute format: " + r;
			}
			break;
		case 115:
			if (r != "s" && r != "ss" && r != ".0" && r != ".00" && r != ".000") throw "bad second format: " + r;
			if (t.u === 0 && (r == "s" || r == "ss")) return W(t.S, r.length);
			if (a >= 2) s = a === 3 ? 1e3: 100;
			else s = a === 1 ? 10 : 1;
			i = Math.round(s * (t.S + t.u));
			if (i >= 60 * s) i = 0;
			if (r === "s") return i === 0 ? "0": "" + i / s;
			n = W(i, 2 + a);
			if (r === "ss") return n.substr(0, 2);
			return "." + n.substr(2, r.length - 1);
		case 90:
			switch (r) {
			case "[h]":
				;
			case "[hh]":
				l = t.D * 24 + t.H;
				break;
			case "[m]":
				;
			case "[mm]":
				l = (t.D * 24 + t.H) * 60 + t.M;
				break;
			case "[s]":
				;
			case "[ss]":
				l = ((t.D * 24 + t.H) * 60 + t.M) * 60 + (a == 0 ? Math.round(t.S + t.u) : t.S);
				break;
			default:
				throw "bad abstime format: " + r;
			}
			o = r.length === 3 ? 1 : 2;
			break;
		case 101:
			l = f;
			o = 1;
			break;
		}
		var c = o > 0 ? W(l, o) : "";
		return c
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
	function ve(e, r, t) {
		var a = r.replace(de, ""),
		n = r.length - a.length;
		return Fe(e, a, t * Math.pow(10, 2 * n)) + yr("%", n)
	}
	function pe(e, r, t) {
		var a = r.length - 1;
		while (r.charCodeAt(a - 1) === 44)--a;
		return Fe(e, r.substr(0, a), t / Math.pow(10, 3 * (r.length - a)))
	}
	function me(e, r) {
		var t;
		var a = e.indexOf("E") - e.indexOf(".") - 1;
		if (e.match(/^#+0.0E\+0$/)) {
			if (r == 0) return "0.0E+0";
			else if (r < 0) return "-" + me(e, -r);
			var n = e.indexOf(".");
			if (n === -1) n = e.indexOf("E");
			var i = Math.floor(Math.log(r) * Math.LOG10E) % n;
			if (i < 0) i += n;
			t = (r / Math.pow(10, i)).toPrecision(a + 1 + (n + i) % n);
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
	var be = /# (\?+)( ?)\/( ?)(\d+)/;
	function ge(e, r, t) {
		var a = parseInt(e[4], 10),
		n = Math.round(r * a),
		i = Math.floor(n / a);
		var s = n - i * a,
		f = a;
		return t + (i === 0 ? "": "" + i) + " " + (s === 0 ? yr(" ", e[1].length + 1 + e[4].length) : z(s, e[1].length) + e[2] + "/" + e[3] + W(f, e[4].length))
	}
	function we(e, r, t) {
		return t + (r === 0 ? "": "" + r) + yr(" ", e[1].length + 2 + e[4].length)
	}
	var ke = /^#*0*\.([0#]+)/;
	var Te = /\).*[0#]/;
	var ye = /\(###\) ###\\?-####/;
	function Ee(e) {
		var r = "",
		t;
		for (var a = 0; a != e.length; ++a) switch (t = e.charCodeAt(a)) {
		case 35:
			break;
		case 63:
			r += " ";
			break;
		case 48:
			r += "0";
			break;
		default:
			r += String.fromCharCode(t);
		}
		return r
	}
	function _e(e, r) {
		var t = Math.pow(10, r);
		return "" + Math.round(e * t) / t
	}
	function Se(e, r) {
		var t = e - Math.floor(e),
		a = Math.pow(10, r);
		if (r < ("" + Math.round(t * a)).length) return 0;
		return Math.round(t * a)
	}
	function xe(e, r) {
		if (r < ("" + Math.round((e - Math.floor(e)) * Math.pow(10, r))).length) {
			return 1
		}
		return 0
	}
	function Ae(e) {
		if (e < 2147483647 && e > -2147483648) return "" + (e >= 0 ? e | 0 : e - 1 | 0);
		return "" + Math.floor(e)
	}
	function Ce(e, r, t) {
		if (e.charCodeAt(0) === 40 && !r.match(Te)) {
			var a = r.replace(/\( */, "").replace(/ \)/, "").replace(/\)/, "");
			if (t >= 0) return Ce("n", a, t);
			return "(" + Ce("n", a, -t) + ")"
		}
		if (r.charCodeAt(r.length - 1) === 44) return pe(e, r, t);
		if (r.indexOf("%") !== -1) return ve(e, r, t);
		if (r.indexOf("E") !== -1) return me(r, t);
		if (r.charCodeAt(0) === 36) return "$" + Ce(e, r.substr(r.charAt(1) == " " ? 2 : 1), t);
		var n;
		var i, s, f, l = Math.abs(t),
		o = t < 0 ? "-": "";
		if (r.match(/^00+$/)) return o + j(l, r.length);
		if (r.match(/^[#?]+$/)) {
			n = j(t, 0);
			if (n === "0") n = "";
			return n.length > r.length ? n: Ee(r.substr(0, r.length - n.length)) + n
		}
		if (i = r.match(be)) return ge(i, l, o);
		if (r.match(/^#+0+$/)) return o + j(l, r.length - r.indexOf("0"));
		if (i = r.match(ke)) {
			n = _e(t, i[1].length).replace(/^([^\.]+)$/, "$1." + Ee(i[1])).replace(/\.$/, "." + Ee(i[1])).replace(/\.(\d*)$/,
			function(e, r) {
				return "." + r + yr("0", Ee(i[1]).length - r.length)
			});
			return r.indexOf("0.") !== -1 ? n: n.replace(/^0\./, ".")
		}
		r = r.replace(/^#+([0.])/, "$1");
		if (i = r.match(/^(0*)\.(#*)$/)) {
			return o + _e(l, i[2].length).replace(/\.(\d*[1-9])0*$/, ".$1").replace(/^(-?\d*)$/, "$1.").replace(/^0\./, i[1].length ? "0.": ".")
		}
		if (i = r.match(/^#{1,3},##0(\.?)$/)) return o + he(j(l, 0));
		if (i = r.match(/^#,##0\.([#0]*0)$/)) {
			return t < 0 ? "-" + Ce(e, r, -t) : he("" + (Math.floor(t) + xe(t, i[1].length))) + "." + W(Se(t, i[1].length), i[1].length)
		}
		if (i = r.match(/^#,#*,#0/)) return Ce(e, r.replace(/^#,#*,/, ""), t);
		if (i = r.match(/^([0#]+)(\\?-([0#]+))+$/)) {
			n = B(Ce(e, r.replace(/[\\-]/g, ""), t));
			s = 0;
			return B(B(r.replace(/\\/g, "")).replace(/[0#]/g,
			function(e) {
				return s < n.length ? n.charAt(s++) : e === "0" ? "0": ""
			}))
		}
		if (r.match(ye)) {
			n = Ce(e, "##########", t);
			return "(" + n.substr(0, 3) + ") " + n.substr(3, 3) + "-" + n.substr(6)
		}
		var c = "";
		if (i = r.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/)) {
			s = Math.min(i[4].length, 7);
			f = re(l, Math.pow(10, s) - 1, false);
			n = "" + o;
			c = Fe("n", i[1], f[1]);
			if (c.charAt(c.length - 1) == " ") c = c.substr(0, c.length - 1) + "0";
			n += c + i[2] + "/" + i[3];
			c = H(f[2], s);
			if (c.length < i[4].length) c = Ee(i[4].substr(i[4].length - c.length)) + c;
			n += c;
			return n
		}
		if (i = r.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/)) {
			s = Math.min(Math.max(i[1].length, i[4].length), 7);
			f = re(l, Math.pow(10, s) - 1, true);
			return o + (f[0] || (f[1] ? "": "0")) + " " + (f[1] ? z(f[1], s) + i[2] + "/" + i[3] + H(f[2], s) : yr(" ", 2 * s + 1 + i[2].length + i[3].length))
		}
		if (i = r.match(/^[#0?]+$/)) {
			n = j(t, 0);
			if (r.length <= n.length) return n;
			return Ee(r.substr(0, r.length - n.length)) + n
		}
		if (i = r.match(/^([#0?]+)\.([#0]+)$/)) {
			n = "" + t.toFixed(Math.min(i[2].length, 10)).replace(/([^0])0+$/, "$1");
			s = n.indexOf(".");
			var u = r.indexOf(".") - s,
			h = r.length - n.length - u;
			return Ee(r.substr(0, u) + n + r.substr(r.length - h))
		}
		if (i = r.match(/^00,000\.([#0]*0)$/)) {
			s = Se(t, i[1].length);
			return t < 0 ? "-" + Ce(e, r, -t) : he(Ae(t)).replace(/^\d,\d{3}$/, "0$&").replace(/^\d*$/,
			function(e) {
				return "00," + (e.length < 3 ? W(0, 3 - e.length) : "") + e
			}) + "." + W(s, i[1].length)
		}
		switch (r) {
		case "###,##0.00":
			return Ce(e, "#,##0.00", t);
		case "###,###":
			;
		case "##,###":
			;
		case "#,###":
			var d = he(j(l, 0));
			return d !== "0" ? o + d: "";
		case "###,###.00":
			return Ce(e, "###,##0.00", t).replace(/^0\./, ".");
		case "#,###.00":
			return Ce(e, "#,##0.00", t).replace(/^0\./, ".");
		default:
			;
		}
		throw new Error("unsupported format |" + r + "|")
	}
	function Re(e, r, t) {
		var a = r.length - 1;
		while (r.charCodeAt(a - 1) === 44)--a;
		return Fe(e, r.substr(0, a), t / Math.pow(10, 3 * (r.length - a)))
	}
	function Oe(e, r, t) {
		var a = r.replace(de, ""),
		n = r.length - a.length;
		return Fe(e, a, t * Math.pow(10, 2 * n)) + yr("%", n)
	}
	function Ie(e, r) {
		var t;
		var a = e.indexOf("E") - e.indexOf(".") - 1;
		if (e.match(/^#+0.0E\+0$/)) {
			if (r == 0) return "0.0E+0";
			else if (r < 0) return "-" + Ie(e, -r);
			var n = e.indexOf(".");
			if (n === -1) n = e.indexOf("E");
			var i = Math.floor(Math.log(r) * Math.LOG10E) % n;
			if (i < 0) i += n;
			t = (r / Math.pow(10, i)).toPrecision(a + 1 + (n + i) % n);
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
	function Ne(e, r, t) {
		if (e.charCodeAt(0) === 40 && !r.match(Te)) {
			var a = r.replace(/\( */, "").replace(/ \)/, "").replace(/\)/, "");
			if (t >= 0) return Ne("n", a, t);
			return "(" + Ne("n", a, -t) + ")"
		}
		if (r.charCodeAt(r.length - 1) === 44) return Re(e, r, t);
		if (r.indexOf("%") !== -1) return Oe(e, r, t);
		if (r.indexOf("E") !== -1) return Ie(r, t);
		if (r.charCodeAt(0) === 36) return "$" + Ne(e, r.substr(r.charAt(1) == " " ? 2 : 1), t);
		var n;
		var i, s, f, l = Math.abs(t),
		o = t < 0 ? "-": "";
		if (r.match(/^00+$/)) return o + W(l, r.length);
		if (r.match(/^[#?]+$/)) {
			n = "" + t;
			if (t === 0) n = "";
			return n.length > r.length ? n: Ee(r.substr(0, r.length - n.length)) + n
		}
		if (i = r.match(be)) return we(i, l, o);
		if (r.match(/^#+0+$/)) return o + W(l, r.length - r.indexOf("0"));
		if (i = r.match(ke)) {
			n = ("" + t).replace(/^([^\.]+)$/, "$1." + Ee(i[1])).replace(/\.$/, "." + Ee(i[1]));
			n = n.replace(/\.(\d*)$/,
			function(e, r) {
				return "." + r + yr("0", Ee(i[1]).length - r.length)
			});
			return r.indexOf("0.") !== -1 ? n: n.replace(/^0\./, ".")
		}
		r = r.replace(/^#+([0.])/, "$1");
		if (i = r.match(/^(0*)\.(#*)$/)) {
			return o + ("" + l).replace(/\.(\d*[1-9])0*$/, ".$1").replace(/^(-?\d*)$/, "$1.").replace(/^0\./, i[1].length ? "0.": ".")
		}
		if (i = r.match(/^#{1,3},##0(\.?)$/)) return o + he("" + l);
		if (i = r.match(/^#,##0\.([#0]*0)$/)) {
			return t < 0 ? "-" + Ne(e, r, -t) : he("" + t) + "." + yr("0", i[1].length)
		}
		if (i = r.match(/^#,#*,#0/)) return Ne(e, r.replace(/^#,#*,/, ""), t);
		if (i = r.match(/^([0#]+)(\\?-([0#]+))+$/)) {
			n = B(Ne(e, r.replace(/[\\-]/g, ""), t));
			s = 0;
			return B(B(r.replace(/\\/g, "")).replace(/[0#]/g,
			function(e) {
				return s < n.length ? n.charAt(s++) : e === "0" ? "0": ""
			}))
		}
		if (r.match(ye)) {
			n = Ne(e, "##########", t);
			return "(" + n.substr(0, 3) + ") " + n.substr(3, 3) + "-" + n.substr(6)
		}
		var c = "";
		if (i = r.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/)) {
			s = Math.min(i[4].length, 7);
			f = re(l, Math.pow(10, s) - 1, false);
			n = "" + o;
			c = Fe("n", i[1], f[1]);
			if (c.charAt(c.length - 1) == " ") c = c.substr(0, c.length - 1) + "0";
			n += c + i[2] + "/" + i[3];
			c = H(f[2], s);
			if (c.length < i[4].length) c = Ee(i[4].substr(i[4].length - c.length)) + c;
			n += c;
			return n
		}
		if (i = r.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/)) {
			s = Math.min(Math.max(i[1].length, i[4].length), 7);
			f = re(l, Math.pow(10, s) - 1, true);
			return o + (f[0] || (f[1] ? "": "0")) + " " + (f[1] ? z(f[1], s) + i[2] + "/" + i[3] + H(f[2], s) : yr(" ", 2 * s + 1 + i[2].length + i[3].length))
		}
		if (i = r.match(/^[#0?]+$/)) {
			n = "" + t;
			if (r.length <= n.length) return n;
			return Ee(r.substr(0, r.length - n.length)) + n
		}
		if (i = r.match(/^([#0]+)\.([#0]+)$/)) {
			n = "" + t.toFixed(Math.min(i[2].length, 10)).replace(/([^0])0+$/, "$1");
			s = n.indexOf(".");
			var u = r.indexOf(".") - s,
			h = r.length - n.length - u;
			return Ee(r.substr(0, u) + n + r.substr(r.length - h))
		}
		if (i = r.match(/^00,000\.([#0]*0)$/)) {
			return t < 0 ? "-" + Ne(e, r, -t) : he("" + t).replace(/^\d,\d{3}$/, "0$&").replace(/^\d*$/,
			function(e) {
				return "00," + (e.length < 3 ? W(0, 3 - e.length) : "") + e
			}) + "." + W(0, i[1].length)
		}
		switch (r) {
		case "###,###":
			;
		case "##,###":
			;
		case "#,###":
			var d = he("" + l);
			return d !== "0" ? o + d: "";
		default:
			if (r.match(/\.[0#?]*$/)) return Ne(e, r.slice(0, r.lastIndexOf(".")), t) + Ee(r.slice(r.lastIndexOf(".")));
		}
		throw new Error("unsupported format |" + r + "|")
	}
	function Fe(e, r, t) {
		return (t | 0) === t ? Ne(e, r, t) : Ce(e, r, t)
	}
	function De(e) {
		var r = [];
		var t = false;
		for (var a = 0,
		n = 0; a < e.length; ++a) switch (e.charCodeAt(a)) {
		case 34:
			t = !t;
			break;
		case 95:
			;
		case 42:
			;
		case 92:
			++a;
			break;
		case 59:
			r[r.length] = e.substr(n, a - n);
			n = a + 1;
		}
		r[r.length] = e.substr(n);
		if (t === true) throw new Error("Format |" + e + "| unterminated string ");
		return r
	}
	var Pe = /\[[HhMmSs\u0E0A\u0E19\u0E17]*\]/;
	function Le(e) {
		var r = 0,
		t = "",
		a = "";
		while (r < e.length) {
			switch (t = e.charAt(r)) {
			case "G":
				if (K(e, r)) r += 6;
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
				;
			case "b":
				if (e.charAt(r + 1) === "1" || e.charAt(r + 1) === "2") return true;
			case "M":
				;
			case "D":
				;
			case "Y":
				;
			case "H":
				;
			case "S":
				;
			case "E":
				;
			case "m":
				;
			case "d":
				;
			case "y":
				;
			case "h":
				;
			case "s":
				;
			case "e":
				;
			case "g":
				return true;
			case "A":
				;
			case "a":
				;
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
				;
			case "0":
				;
			case "#":
				while (r < e.length && ("0#?.,E+-%".indexOf(t = e.charAt(++r)) > -1 || t == "\\" && e.charAt(r + 1) == "-" && "0#".indexOf(e.charAt(r + 2)) > -1)) {}
				break;
			case "?":
				while (e.charAt(++r) === t) {}
				break;
			case "*":
				++r;
				if (e.charAt(r) == " " || e.charAt(r) == "*")++r;
				break;
			case "(":
				;
			case ")":
				++r;
				break;
			case "1":
				;
			case "2":
				;
			case "3":
				;
			case "4":
				;
			case "5":
				;
			case "6":
				;
			case "7":
				;
			case "8":
				;
			case "9":
				while (r < e.length && "0123456789".indexOf(e.charAt(++r)) > -1) {}
				break;
			case " ":
				++r;
				break;
			default:
				++r;
				break;
			}
		}
		return false
	}
	function Me(e, r, t, a) {
		var n = [],
		i = "",
		s = 0,
		f = "",
		l = "t",
		o,
		c,
		u;
		var h = "H";
		while (s < e.length) {
			switch (f = e.charAt(s)) {
			case "G":
				if (!K(e, s)) throw new Error("unrecognized character " + f + " in " + e);
				n[n.length] = {
					t: "G",
					v: "General"
				};
				s += 7;
				break;
			case '"':
				for (i = ""; (u = e.charCodeAt(++s)) !== 34 && s < e.length;) i += String.fromCharCode(u);
				n[n.length] = {
					t: "t",
					v: i
				}; ++s;
				break;
			case "\\":
				var d = e.charAt(++s),
				v = d === "(" || d === ")" ? d: "t";
				n[n.length] = {
					t: v,
					v: d
				}; ++s;
				break;
			case "_":
				n[n.length] = {
					t: "t",
					v: " "
				};
				s += 2;
				break;
			case "@":
				n[n.length] = {
					t: "T",
					v: r
				}; ++s;
				break;
			case "B":
				;
			case "b":
				if (e.charAt(s + 1) === "1" || e.charAt(s + 1) === "2") {
					if (o == null) {
						o = ae(r, t, e.charAt(s + 1) === "2");
						if (o == null) return ""
					}
					n[n.length] = {
						t: "X",
						v: e.substr(s, 2)
					};
					l = f;
					s += 2;
					break
				};
			case "M":
				;
			case "D":
				;
			case "Y":
				;
			case "H":
				;
			case "S":
				;
			case "E":
				f = f.toLowerCase();
			case "m":
				;
			case "d":
				;
			case "y":
				;
			case "h":
				;
			case "s":
				;
			case "e":
				;
			case "g":
				if (r < 0) return "";
				if (o == null) {
					o = ae(r, t);
					if (o == null) return ""
				}
				i = f;
				while (++s < e.length && e.charAt(s).toLowerCase() === f) i += f;
				if (f === "m" && l.toLowerCase() === "h") f = "M";
				if (f === "h") f = h;
				n[n.length] = {
					t: f,
					v: i
				};
				l = f;
				break;
			case "A":
				;
			case "a":
				;
			case "上":
				var p = {
					t: f,
					v: f
				};
				if (o == null) o = ae(r, t);
				if (e.substr(s, 3).toUpperCase() === "A/P") {
					if (o != null) p.v = o.H >= 12 ? e.charAt(s + 2) : f;
					p.t = "T";
					h = "h";
					s += 3
				} else if (e.substr(s, 5).toUpperCase() === "AM/PM") {
					if (o != null) p.v = o.H >= 12 ? "PM": "AM";
					p.t = "T";
					s += 5;
					h = "h"
				} else if (e.substr(s, 5).toUpperCase() === "上午/下午") {
					if (o != null) p.v = o.H >= 12 ? "下午": "上午";
					p.t = "T";
					s += 5;
					h = "h"
				} else {
					p.t = "t"; ++s
				}
				if (o == null && p.t === "T") return "";
				n[n.length] = p;
				l = f;
				break;
			case "[":
				i = f;
				while (e.charAt(s++) !== "]" && s < e.length) i += e.charAt(s);
				if (i.slice(-1) !== "]") throw 'unterminated "[" block: |' + i + "|";
				if (i.match(Pe)) {
					if (o == null) {
						o = ae(r, t);
						if (o == null) return ""
					}
					n[n.length] = {
						t: "Z",
						v: i.toLowerCase()
					};
					l = i.charAt(1)
				} else if (i.indexOf("$") > -1) {
					i = (i.match(/\$([^-\[\]]*)/) || [])[1] || "$";
					if (!Le(e)) n[n.length] = {
						t: "t",
						v: i
					}
				}
				break;
			case ".":
				if (o != null) {
					i = f;
					while (++s < e.length && (f = e.charAt(s)) === "0") i += f;
					n[n.length] = {
						t: "s",
						v: i
					};
					break
				};
			case "0":
				;
			case "#":
				i = f;
				while (++s < e.length && "0#?.,E+-%".indexOf(f = e.charAt(s)) > -1) i += f;
				n[n.length] = {
					t: "n",
					v: i
				};
				break;
			case "?":
				i = f;
				while (e.charAt(++s) === f) i += f;
				n[n.length] = {
					t: f,
					v: i
				};
				l = f;
				break;
			case "*":
				++s;
				if (e.charAt(s) == " " || e.charAt(s) == "*")++s;
				break;
			case "(":
				;
			case ")":
				n[n.length] = {
					t: a === 1 ? "t": f,
					v: f
				}; ++s;
				break;
			case "1":
				;
			case "2":
				;
			case "3":
				;
			case "4":
				;
			case "5":
				;
			case "6":
				;
			case "7":
				;
			case "8":
				;
			case "9":
				i = f;
				while (s < e.length && "0123456789".indexOf(e.charAt(++s)) > -1) i += e.charAt(s);
				n[n.length] = {
					t: "D",
					v: i
				};
				break;
			case " ":
				n[n.length] = {
					t: f,
					v: f
				}; ++s;
				break;
			case "$":
				n[n.length] = {
					t: "t",
					v: "$"
				}; ++s;
				break;
			default:
				if (",$-+/():!^&'~{}<>=€acfijklopqrtuvwxzP".indexOf(f) === -1) throw new Error("unrecognized character " + f + " in " + e);
				n[n.length] = {
					t: "t",
					v: f
				}; ++s;
				break;
			}
		}
		var m = 0,
		b = 0,
		g;
		for (s = n.length - 1, l = "t"; s >= 0; --s) {
			switch (n[s].t) {
			case "h":
				;
			case "H":
				n[s].t = h;
				l = "h";
				if (m < 1) m = 1;
				break;
			case "s":
				if (g = n[s].v.match(/\.0+$/)) {
					b = Math.max(b, g[0].length - 1);
					m = 4
				}
				if (m < 3) m = 3;
			case "d":
				;
			case "y":
				;
			case "e":
				l = n[s].t;
				break;
			case "M":
				l = n[s].t;
				if (m < 2) m = 2;
				break;
			case "m":
				if (l === "s") {
					n[s].t = "M";
					if (m < 2) m = 2
				}
				break;
			case "X":
				break;
			case "Z":
				if (m < 1 && n[s].v.match(/[Hh]/)) m = 1;
				if (m < 2 && n[s].v.match(/[Mm]/)) m = 2;
				if (m < 3 && n[s].v.match(/[Ss]/)) m = 3;
			}
		}
		var w;
		switch (m) {
		case 0:
			break;
		case 1:
			;
		case 2:
			;
		case 3:
			if (o.u >= .5) {
				o.u = 0; ++o.S
			}
			if (o.S >= 60) {
				o.S = 0; ++o.M
			}
			if (o.M >= 60) {
				o.M = 0; ++o.H
			}
			if (o.H >= 24) {
				o.H = 0; ++o.D;
				w = ae(o.D);
				w.u = o.u;
				w.S = o.S;
				w.M = o.M;
				w.H = o.H;
				o = w
			}
			break;
		case 4:
			switch (b) {
			case 1:
				o.u = Math.round(o.u * 10) / 10;
				break;
			case 2:
				o.u = Math.round(o.u * 100) / 100;
				break;
			case 3:
				o.u = Math.round(o.u * 1e3) / 1e3;
				break;
			}
			if (o.u >= 1) {
				o.u = 0; ++o.S
			}
			if (o.S >= 60) {
				o.S = 0; ++o.M
			}
			if (o.M >= 60) {
				o.M = 0; ++o.H
			}
			if (o.H >= 24) {
				o.H = 0; ++o.D;
				w = ae(o.D);
				w.u = o.u;
				w.S = o.S;
				w.M = o.M;
				w.H = o.H;
				o = w
			}
			break;
		}
		var k = "",
		T;
		for (s = 0; s < n.length; ++s) {
			switch (n[s].t) {
			case "t":
				;
			case "T":
				;
			case " ":
				;
			case "D":
				break;
			case "X":
				n[s].v = "";
				n[s].t = ";";
				break;
			case "d":
				;
			case "m":
				;
			case "y":
				;
			case "h":
				;
			case "H":
				;
			case "M":
				;
			case "s":
				;
			case "e":
				;
			case "b":
				;
			case "Z":
				n[s].v = ue(n[s].t.charCodeAt(0), n[s].v, o, b);
				n[s].t = "t";
				break;
			case "n":
				;
			case "?":
				T = s + 1;
				while (n[T] != null && ((f = n[T].t) === "?" || f === "D" || (f === " " || f === "t") && n[T + 1] != null && (n[T + 1].t === "?" || n[T + 1].t === "t" && n[T + 1].v === "/") || n[s].t === "(" && (f === " " || f === "n" || f === ")") || f === "t" && (n[T].v === "/" || n[T].v === " " && n[T + 1] != null && n[T + 1].t == "?"))) {
					n[s].v += n[T].v;
					n[T] = {
						v: "",
						t: ";"
					}; ++T
				}
				k += n[s].v;
				s = T - 1;
				break;
			case "G":
				n[s].t = "t";
				n[s].v = oe(r, t);
				break;
			}
		}
		var y = "",
		E, _;
		if (k.length > 0) {
			if (k.charCodeAt(0) == 40) {
				E = r < 0 && k.charCodeAt(0) === 45 ? -r: r;
				_ = Fe("n", k, E)
			} else {
				E = r < 0 && a > 1 ? -r: r;
				_ = Fe("n", k, E);
				if (E < 0 && n[0] && n[0].t == "t") {
					_ = _.substr(1);
					n[0].v = "-" + n[0].v
				}
			}
			T = _.length - 1;
			var S = n.length;
			for (s = 0; s < n.length; ++s) if (n[s] != null && n[s].t != "t" && n[s].v.indexOf(".") > -1) {
				S = s;
				break
			}
			var x = n.length;
			if (S === n.length && _.indexOf("E") === -1) {
				for (s = n.length - 1; s >= 0; --s) {
					if (n[s] == null || "n?".indexOf(n[s].t) === -1) continue;
					if (T >= n[s].v.length - 1) {
						T -= n[s].v.length;
						n[s].v = _.substr(T + 1, n[s].v.length)
					} else if (T < 0) n[s].v = "";
					else {
						n[s].v = _.substr(0, T + 1);
						T = -1
					}
					n[s].t = "t";
					x = s
				}
				if (T >= 0 && x < n.length) n[x].v = _.substr(0, T + 1) + n[x].v
			} else if (S !== n.length && _.indexOf("E") === -1) {
				T = _.indexOf(".") - 1;
				for (s = S; s >= 0; --s) {
					if (n[s] == null || "n?".indexOf(n[s].t) === -1) continue;
					c = n[s].v.indexOf(".") > -1 && s === S ? n[s].v.indexOf(".") - 1 : n[s].v.length - 1;
					y = n[s].v.substr(c + 1);
					for (; c >= 0; --c) {
						if (T >= 0 && (n[s].v.charAt(c) === "0" || n[s].v.charAt(c) === "#")) y = _.charAt(T--) + y
					}
					n[s].v = y;
					n[s].t = "t";
					x = s
				}
				if (T >= 0 && x < n.length) n[x].v = _.substr(0, T + 1) + n[x].v;
				T = _.indexOf(".") + 1;
				for (s = S; s < n.length; ++s) {
					if (n[s] == null || "n?(".indexOf(n[s].t) === -1 && s !== S) continue;
					c = n[s].v.indexOf(".") > -1 && s === S ? n[s].v.indexOf(".") + 1 : 0;
					y = n[s].v.substr(0, c);
					for (; c < n[s].v.length; ++c) {
						if (T < _.length) y += _.charAt(T++)
					}
					n[s].v = y;
					n[s].t = "t";
					x = s
				}
			}
		}
		for (s = 0; s < n.length; ++s) if (n[s] != null && "n?".indexOf(n[s].t) > -1) {
			E = a > 1 && r < 0 && s > 0 && n[s - 1].v === "-" ? -r: r;
			n[s].v = Fe(n[s].t, n[s].v, E);
			n[s].t = "t"
		}
		var A = "";
		for (s = 0; s !== n.length; ++s) if (n[s] != null) A += n[s].v;
		return A
	}
	var Ue = /\[(=|>[=]?|<[>=]?)(-?\d+(?:\.\d*)?)\]/;
	function Be(e, r) {
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
			break;
		}
		return false
	}
	function We(e, r) {
		var t = De(e);
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
			break;
		}
		var i = r > 0 ? t[0] : r < 0 ? t[1] : t[2];
		if (t[0].indexOf("[") === -1 && t[1].indexOf("[") === -1) return [a, i];
		if (t[0].match(/\[[=<>]/) != null || t[1].match(/\[[=<>]/) != null) {
			var s = t[0].match(Ue);
			var f = t[1].match(Ue);
			return Be(r, s) ? [a, t[0]] : Be(r, f) ? [a, t[1]] : [a, t[s != null && f != null ? 2 : 1]]
		}
		return [a, i]
	}
	function ze(e, r, t) {
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
			if (a == null) a = t.table && t.table[Q[e]] || q[Q[e]];
			if (a == null) a = ee[e] || "General";
			break;
		}
		if (K(a, 0)) return oe(r, t);
		if (r instanceof Date) r = dr(r, t.date1904);
		var n = We(a, r);
		if (K(n[1])) return oe(r, t);
		if (r === true) r = "TRUE";
		else if (r === false) r = "FALSE";
		else if (r === "" || r == null) return "";
		return Me(n[1], r, t, n[0])
	}
	function He(e, r) {
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
	function Ve(e) {
		for (var r = 0; r != 392; ++r) if (e[r] !== undefined) He(e[r], r)
	}
	function $e() {
		q = J()
	}
	var Xe = {
		format: ze,
		load: He,
		_table: q,
		load_table: Ve,
		parse_date_code: ae,
		is_date: Le,
		get_table: function hT() {
			return Xe._table = q
		}
	};
	var Ge = {
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
		81 : "mmss.0"
	};
	var je = /[dD]+|[mM]+|[yYeE]+|[Hh]+|[Ss]+/g;
	function Ke(e) {
		var r = typeof e == "number" ? q[e] : e;
		r = r.replace(je, "(\\d+)");
		je.lastIndex = 0;
		return new RegExp("^" + r + "$")
	}
	function Ye(e, r, t) {
		var a = -1,
		n = -1,
		i = -1,
		s = -1,
		f = -1,
		l = -1; (r.match(je) || []).forEach(function(e, r) {
			var o = parseInt(t[r + 1], 10);
			switch (e.toLowerCase().charAt(0)) {
			case "y":
				a = o;
				break;
			case "d":
				i = o;
				break;
			case "h":
				s = o;
				break;
			case "s":
				l = o;
				break;
			case "m":
				if (s >= 0) f = o;
				else n = o;
				break;
			}
		});
		je.lastIndex = 0;
		if (l >= 0 && f == -1 && n >= 0) {
			f = n;
			n = -1
		}
		var o = ("" + (a >= 0 ? a: (new Date).getFullYear())).slice(-4) + "-" + ("00" + (n >= 1 ? n: 1)).slice(-2) + "-" + ("00" + (i >= 1 ? i: 1)).slice(-2);
		if (o.length == 7) o = "0" + o;
		if (o.length == 8) o = "20" + o;
		var c = ("00" + (s >= 0 ? s: 0)).slice(-2) + ":" + ("00" + (f >= 0 ? f: 0)).slice(-2) + ":" + ("00" + (l >= 0 ? l: 0)).slice(-2);
		if (s == -1 && f == -1 && l == -1) return o;
		if (a == -1 && n == -1 && i == -1) return c;
		return o + "T" + c
	}
	var Ze = {
		"d.m": "d\\.m"
	};
	function Je(e, r) {
		return He(Ze[e] || e, r)
	}
	var qe = function() {
		var e = {};
		e.version = "1.2.0";
		function r() {
			var e = 0,
			r = new Array(256);
			for (var t = 0; t != 256; ++t) {
				e = t;
				e = e & 1 ? -306674912 ^ e >>> 1 : e >>> 1;
				e = e & 1 ? -306674912 ^ e >>> 1 : e >>> 1;
				e = e & 1 ? -306674912 ^ e >>> 1 : e >>> 1;
				e = e & 1 ? -306674912 ^ e >>> 1 : e >>> 1;
				e = e & 1 ? -306674912 ^ e >>> 1 : e >>> 1;
				e = e & 1 ? -306674912 ^ e >>> 1 : e >>> 1;
				e = e & 1 ? -306674912 ^ e >>> 1 : e >>> 1;
				e = e & 1 ? -306674912 ^ e >>> 1 : e >>> 1;
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
				for (r = 256 + a; r < 4096; r += 256) t = n[r] = t >>> 8 ^ e[t & 255]
			}
			var i = [];
			for (a = 1; a != 16; ++a) i[a - 1] = typeof Int32Array !== "undefined" && typeof n.subarray == "function" ? n.subarray(a * 256, a * 256 + 256) : n.slice(a * 256, a * 256 + 256);
			return i
		}
		var n = a(t);
		var i = n[0],
		s = n[1],
		f = n[2],
		l = n[3],
		o = n[4];
		var c = n[5],
		u = n[6],
		h = n[7],
		d = n[8],
		v = n[9];
		var p = n[10],
		m = n[11],
		b = n[12],
		g = n[13],
		w = n[14];
		function k(e, r) {
			var a = r ^ -1;
			for (var n = 0,
			i = e.length; n < i;) a = a >>> 8 ^ t[(a ^ e.charCodeAt(n++)) & 255];
			return~a
		}
		function T(e, r) {
			var a = r ^ -1,
			n = e.length - 15,
			k = 0;
			for (; k < n;) a = w[e[k++] ^ a & 255] ^ g[e[k++] ^ a >> 8 & 255] ^ b[e[k++] ^ a >> 16 & 255] ^ m[e[k++] ^ a >>> 24] ^ p[e[k++]] ^ v[e[k++]] ^ d[e[k++]] ^ h[e[k++]] ^ u[e[k++]] ^ c[e[k++]] ^ o[e[k++]] ^ l[e[k++]] ^ f[e[k++]] ^ s[e[k++]] ^ i[e[k++]] ^ t[e[k++]];
			n += 15;
			while (k < n) a = a >>> 8 ^ t[(a ^ e[k++]) & 255];
			return~a
		}
		function y(e, r) {
			var a = r ^ -1;
			for (var n = 0,
			i = e.length,
			s = 0,
			f = 0; n < i;) {
				s = e.charCodeAt(n++);
				if (s < 128) {
					a = a >>> 8 ^ t[(a ^ s) & 255]
				} else if (s < 2048) {
					a = a >>> 8 ^ t[(a ^ (192 | s >> 6 & 31)) & 255];
					a = a >>> 8 ^ t[(a ^ (128 | s & 63)) & 255]
				} else if (s >= 55296 && s < 57344) {
					s = (s & 1023) + 64;
					f = e.charCodeAt(n++) & 1023;
					a = a >>> 8 ^ t[(a ^ (240 | s >> 8 & 7)) & 255];
					a = a >>> 8 ^ t[(a ^ (128 | s >> 2 & 63)) & 255];
					a = a >>> 8 ^ t[(a ^ (128 | f >> 6 & 15 | (s & 3) << 4)) & 255];
					a = a >>> 8 ^ t[(a ^ (128 | f & 63)) & 255]
				} else {
					a = a >>> 8 ^ t[(a ^ (224 | s >> 12 & 15)) & 255];
					a = a >>> 8 ^ t[(a ^ (128 | s >> 6 & 63)) & 255];
					a = a >>> 8 ^ t[(a ^ (128 | s & 63)) & 255]
				}
			}
			return~a
		}
		e.table = t;
		e.bstr = k;
		e.buf = T;
		e.str = y;
		return e
	} ();
	var Qe = function dT() {
		var e = {};
		e.version = "1.2.2";
		function r(e, r) {
			var t = e.split("/"),
			a = r.split("/");
			for (var n = 0,
			i = 0,
			s = Math.min(t.length, a.length); n < s; ++n) {
				if (i = t[n].length - a[n].length) return i;
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
			t = t << 6 | r.getMinutes();
			t = t << 5 | r.getSeconds() >>> 1;
			e._W(2, t);
			var a = r.getFullYear() - 1980;
			a = a << 4 | r.getMonth() + 1;
			a = a << 5 | r.getDate();
			e._W(2, a)
		}
		function i(e) {
			var r = e._R(2) & 65535;
			var t = e._R(2) & 65535;
			var a = new Date;
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
			var f = r & 63;
			r >>>= 6;
			a.setHours(r);
			a.setMinutes(f);
			a.setSeconds(s << 1);
			return a
		}
		function s(e) {
			Ta(e, 0);
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
						var f = e._R(4),
						l = e._R(4);
						s.usz = l * Math.pow(2, 32) + f;
						f = e._R(4);
						l = e._R(4);
						s.csz = l * Math.pow(2, 32) + f
					}
					break;
				}
				e.l = i;
				r[a] = s
			}
			return r
		}
		var f;
		function l() {
			return f || (f = er)
		}
		function o(e, r) {
			if (e[0] == 80 && e[1] == 75) return Ie(e, r);
			if ((e[0] | 32) == 109 && (e[1] | 32) == 105) return We(e, r);
			if (e.length < 512) throw new Error("CFB file size " + e.length + " < 512");
			var t = 3;
			var a = 512;
			var n = 0;
			var i = 0;
			var s = 0;
			var f = 0;
			var l = 0;
			var o = [];
			var v = e.slice(0, 512);
			Ta(v, 0);
			var m = c(v);
			t = m[0];
			switch (t) {
			case 3:
				a = 512;
				break;
			case 4:
				a = 4096;
				break;
			case 0:
				if (m[1] == 0) return Ie(e, r);
			default:
				throw new Error("Major Version: Expected 3 or 4 saw " + t);
			}
			if (a !== 512) {
				v = e.slice(0, a);
				Ta(v, 28)
			}
			var w = e.slice(0, a);
			u(v, t);
			var k = v._R(4, "i");
			if (t === 3 && k !== 0) throw new Error("# Directory Sectors: Expected 0 saw " + k);
			v.l += 4;
			s = v._R(4, "i");
			v.l += 4;
			v.chk("00100000", "Mini Stream Cutoff Size: ");
			f = v._R(4, "i");
			n = v._R(4, "i");
			l = v._R(4, "i");
			i = v._R(4, "i");
			for (var T = -1,
			y = 0; y < 109; ++y) {
				T = v._R(4, "i");
				if (T < 0) break;
				o[y] = T
			}
			var E = h(e, a);
			p(l, i, E, a, o);
			var _ = b(E, s, o, a);
			if (s < _.length) _[s].name = "!Directory";
			if (n > 0 && f !== L) _[f].name = "!MiniFAT";
			_[o[0]].name = "!FAT";
			_.fat_addrs = o;
			_.ssz = a;
			var S = {},
			x = [],
			A = [],
			C = [];
			g(s, _, E, x, n, S, A, f);
			d(A, C, x);
			x.shift();
			var R = {
				FileIndex: A,
				FullPaths: C
			};
			if (r && r.raw) R.raw = {
				header: w,
				sectors: E
			};
			return R
		}
		function c(e) {
			if (e[e.l] == 80 && e[e.l + 1] == 75) return [0, 0];
			e.chk(B, "Header Signature: ");
			e.l += 16;
			var r = e._R(2, "u");
			return [e._R(2, "u"), r]
		}
		function u(e, r) {
			var t = 9;
			e.l += 2;
			switch (t = e._R(2)) {
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
			f = 0,
			l = t.length;
			var o = [],
			c = [];
			for (; a < l; ++a) {
				o[a] = c[a] = a;
				r[a] = t[a]
			}
			for (; f < c.length; ++f) {
				a = c[f];
				n = e[a].L;
				i = e[a].R;
				s = e[a].C;
				if (o[a] === a) {
					if (n !== -1 && o[n] !== n) o[a] = o[n];
					if (i !== -1 && o[i] !== i) o[a] = o[i]
				}
				if (s !== -1) o[s] = a;
				if (n !== -1 && a != o[a]) {
					o[n] = o[a];
					if (c.lastIndexOf(n) < f) c.push(n)
				}
				if (i !== -1 && a != o[a]) {
					o[i] = o[a];
					if (c.lastIndexOf(i) < f) c.push(i)
				}
			}
			for (a = 1; a < l; ++a) if (o[a] === a) {
				if (i !== -1 && o[i] !== i) o[a] = o[i];
				else if (n !== -1 && o[n] !== n) o[a] = o[n]
			}
			for (a = 1; a < l; ++a) {
				if (e[a].type === 0) continue;
				f = a;
				if (f != o[f]) do {
					f = o[f];
					r[a] = r[f] + "/" + r[a]
				} while ( f !== 0 && - 1 !== o [ f ] && f != o[f]);
				o[a] = -1
			}
			r[0] += "/";
			for (a = 1; a < l; ++a) {
				if (e[a].type !== 2) r[a] += "/"
			}
		}
		function v(e, r, t) {
			var a = e.start,
			n = e.size;
			var i = [];
			var s = a;
			while (t && n > 0 && s >= 0) {
				i.push(r.slice(s * D, s * D + D));
				n -= D;
				s = da(t, s * 4)
			}
			if (i.length === 0) return Ea(0);
			return P(i).slice(0, e.size)
		}
		function p(e, r, t, a, n) {
			var i = L;
			if (e === L) {
				if (r !== 0) throw new Error("DIFAT chain shorter than expected")
			} else if (e !== -1) {
				var s = t[e],
				f = (a >>> 2) - 1;
				if (!s) return;
				for (var l = 0; l < f; ++l) {
					if ((i = da(s, l * 4)) === L) break;
					n.push(i)
				}
				if (r >= 1) p(da(s, a - 4), r - 1, t, a, n)
			}
		}
		function m(e, r, t, a, n) {
			var i = [],
			s = [];
			if (!n) n = [];
			var f = a - 1,
			l = 0,
			o = 0;
			for (l = r; l >= 0;) {
				n[l] = true;
				i[i.length] = l;
				s.push(e[l]);
				var c = t[Math.floor(l * 4 / a)];
				o = l * 4 & f;
				if (a < 4 + o) throw new Error("FAT boundary crossed: " + l + " 4 " + a);
				if (!e[c]) break;
				l = da(e[c], o)
			}
			return {
				nodes: i,
				data: Ht([s])
			}
		}
		function b(e, r, t, a) {
			var n = e.length,
			i = [];
			var s = [],
			f = [],
			l = [];
			var o = a - 1,
			c = 0,
			u = 0,
			h = 0,
			d = 0;
			for (c = 0; c < n; ++c) {
				f = [];
				h = c + r;
				if (h >= n) h -= n;
				if (s[h]) continue;
				l = [];
				var v = [];
				for (u = h; u >= 0;) {
					v[u] = true;
					s[u] = true;
					f[f.length] = u;
					l.push(e[u]);
					var p = t[Math.floor(u * 4 / a)];
					d = u * 4 & o;
					if (a < 4 + d) throw new Error("FAT boundary crossed: " + u + " 4 " + a);
					if (!e[p]) break;
					u = da(e[p], d);
					if (v[u]) break
				}
				i[h] = {
					nodes: f,
					data: Ht([l])
				}
			}
			return i
		}
		function g(e, r, t, a, n, i, s, f) {
			var l = 0,
			o = a.length ? 2 : 0;
			var c = r[e].data;
			var u = 0,
			h = 0,
			d;
			for (; u < c.length; u += 128) {
				var p = c.slice(u, u + 128);
				Ta(p, 64);
				h = p._R(2);
				d = $t(p, 0, h - o);
				a.push(d);
				var b = {
					name: d,
					type: p._R(1),
					color: p._R(1),
					L: p._R(4, "i"),
					R: p._R(4, "i"),
					C: p._R(4, "i"),
					clsid: p._R(16),
					state: p._R(4, "i"),
					start: 0,
					size: 0
				};
				var g = p._R(2) + p._R(2) + p._R(2) + p._R(2);
				if (g !== 0) b.ct = w(p, p.l - 8);
				var k = p._R(2) + p._R(2) + p._R(2) + p._R(2);
				if (k !== 0) b.mt = w(p, p.l - 8);
				b.start = p._R(4, "i");
				b.size = p._R(4, "i");
				if (b.size < 0 && b.start < 0) {
					b.size = b.type = 0;
					b.start = L;
					b.name = ""
				}
				if (b.type === 5) {
					l = b.start;
					if (n > 0 && l !== L) r[l].name = "!StreamData"
				} else if (b.size >= 4096) {
					b.storage = "fat";
					if (r[b.start] === undefined) r[b.start] = m(t, b.start, r.fat_addrs, r.ssz);
					r[b.start].name = b.name;
					b.content = r[b.start].data.slice(0, b.size)
				} else {
					b.storage = "minifat";
					if (b.size < 0) b.size = 0;
					else if (l !== L && b.start !== L && r[l]) {
						b.content = v(b, r[l].data, (r[f] || {}).data)
					}
				}
				if (b.content) Ta(b.content, 0);
				i[d] = b;
				s.push(b)
			}
		}
		function w(e, r) {
			return new Date((ha(e, r + 4) / 1e7 * Math.pow(2, 32) + ha(e, r) / 1e7 - 11644473600) * 1e3)
		}
		function k(e, r) {
			l();
			return o(f.readFileSync(e), r)
		}
		function y(e, r) {
			var t = r && r.type;
			if (!t) {
				if (S && Buffer.isBuffer(e)) t = "buffer"
			}
			switch (t || "base64") {
			case "file":
				return k(e, r);
			case "base64":
				return o(O(_(e)), r);
			case "binary":
				return o(O(e), r);
			}
			return o(e, r)
		}
		function E(e, r) {
			var t = r || {},
			a = t.root || "Root Entry";
			if (!e.FullPaths) e.FullPaths = [];
			if (!e.FileIndex) e.FileIndex = [];
			if (e.FullPaths.length !== e.FileIndex.length) throw new Error("inconsistent CFB structure");
			if (e.FullPaths.length === 0) {
				e.FullPaths[0] = a + "/";
				e.FileIndex[0] = {
					name: a,
					type: 5
				}
			}
			if (t.CLSID) e.FileIndex[0].clsid = t.CLSID;
			A(e)
		}
		function A(e) {
			var r = "Sh33tJ5";
			if (Qe.find(e, "/" + r)) return;
			var t = Ea(4);
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
				C: 69
			});
			e.FullPaths.push(e.FullPaths[0] + r);
			I(e)
		}
		function I(e, n) {
			E(e);
			var i = false,
			s = false;
			for (var f = e.FullPaths.length - 1; f >= 0; --f) {
				var l = e.FileIndex[f];
				switch (l.type) {
				case 0:
					if (s) i = true;
					else {
						e.FileIndex.pop();
						e.FullPaths.pop()
					}
					break;
				case 1:
					;
				case 2:
					;
				case 5:
					s = true;
					if (isNaN(l.R * l.L * l.C)) i = true;
					if (l.R > -1 && l.L > -1 && l.R == l.L) i = true;
					break;
				default:
					i = true;
					break;
				}
			}
			if (!i && !n) return;
			var o = new Date(1987, 1, 19),
			c = 0;
			var u = Object.create ? Object.create(null) : {};
			var h = [];
			for (f = 0; f < e.FullPaths.length; ++f) {
				u[e.FullPaths[f]] = true;
				if (e.FileIndex[f].type === 0) continue;
				h.push([e.FullPaths[f], e.FileIndex[f]])
			}
			for (f = 0; f < h.length; ++f) {
				var d = t(h[f][0]);
				s = u[d];
				while (!s) {
					while (t(d) && !u[t(d)]) d = t(d);
					h.push([d, {
						name: a(d).replace("/", ""),
						type: 1,
						clsid: z,
						ct: o,
						mt: o,
						content: null
					}]);
					u[d] = true;
					d = t(h[f][0]);
					s = u[d]
				}
			}
			h.sort(function(e, t) {
				return r(e[0], t[0])
			});
			e.FullPaths = [];
			e.FileIndex = [];
			for (f = 0; f < h.length; ++f) {
				e.FullPaths[f] = h[f][0];
				e.FileIndex[f] = h[f][1]
			}
			for (f = 0; f < h.length; ++f) {
				var v = e.FileIndex[f];
				var p = e.FullPaths[f];
				v.name = a(p).replace("/", "");
				v.L = v.R = v.C = -(v.color = 1);
				v.size = v.content ? v.content.length: 0;
				v.start = 0;
				v.clsid = v.clsid || z;
				if (f === 0) {
					v.C = h.length > 1 ? 1 : -1;
					v.size = 0;
					v.type = 5
				} else if (p.slice(-1) == "/") {
					for (c = f + 1; c < h.length; ++c) if (t(e.FullPaths[c]) == p) break;
					v.C = c >= h.length ? -1 : c;
					for (c = f + 1; c < h.length; ++c) if (t(e.FullPaths[c]) == t(p)) break;
					v.R = c >= h.length ? -1 : c;
					v.type = 1
				} else {
					if (t(e.FullPaths[f + 1] || "") == t(p)) v.R = f + 1;
					v.type = 2
				}
			}
		}
		function N(e, r) {
			var t = r || {};
			if (t.fileType == "mad") return ze(e, t);
			I(e);
			switch (t.fileType) {
			case "zip":
				return Fe(e, t);
			}
			var a = function(e) {
				var r = 0,
				t = 0;
				for (var a = 0; a < e.FileIndex.length; ++a) {
					var n = e.FileIndex[a];
					if (!n.content) continue;
					var i = n.content.length;
					if (i > 0) {
						if (i < 4096) r += i + 63 >> 6;
						else t += i + 511 >> 9
					}
				}
				var s = e.FullPaths.length + 3 >> 2;
				var f = r + 7 >> 3;
				var l = r + 127 >> 7;
				var o = f + t + s + l;
				var c = o + 127 >> 7;
				var u = c <= 109 ? 0 : Math.ceil((c - 109) / 127);
				while (o + c + u + 127 >> 7 > c) u = ++c <= 109 ? 0 : Math.ceil((c - 109) / 127);
				var h = [1, u, c, l, s, t, r, 0];
				e.FileIndex[0].size = r << 6;
				h[7] = (e.FileIndex[0].start = h[0] + h[1] + h[2] + h[3] + h[4] + h[5]) + (h[6] + 7 >> 3);
				return h
			} (e);
			var n = Ea(a[7] << 9);
			var i = 0,
			s = 0; {
				for (i = 0; i < 8; ++i) n._W(1, W[i]);
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
				n._W(4, a[3] ? a[0] + a[1] + a[2] - 1 : L);
				n._W(4, a[3]);
				n._W(-4, a[1] ? a[0] - 1 : L);
				n._W(4, a[1]);
				for (i = 0; i < 109; ++i) n._W(-4, i < a[2] ? a[1] + i: -1)
			}
			if (a[1]) {
				for (s = 0; s < a[1]; ++s) {
					for (; i < 236 + s * 127; ++i) n._W(-4, i < a[2] ? a[1] + i: -1);
					n._W(-4, s === a[1] - 1 ? L: s + 1)
				}
			}
			var f = function(e) {
				for (s += e; i < s - 1; ++i) n._W(-4, i + 1);
				if (e) {++i;
					n._W(-4, L)
				}
			};
			s = i = 0;
			for (s += a[1]; i < s; ++i) n._W(-4, H.DIFSECT);
			for (s += a[2]; i < s; ++i) n._W(-4, H.FATSECT);
			f(a[3]);
			f(a[4]);
			var l = 0,
			o = 0;
			var c = e.FileIndex[0];
			for (; l < e.FileIndex.length; ++l) {
				c = e.FileIndex[l];
				if (!c.content) continue;
				o = c.content.length;
				if (o < 4096) continue;
				c.start = s;
				f(o + 511 >> 9)
			}
			f(a[6] + 7 >> 3);
			while (n.l & 511) n._W(-4, H.ENDOFCHAIN);
			s = i = 0;
			for (l = 0; l < e.FileIndex.length; ++l) {
				c = e.FileIndex[l];
				if (!c.content) continue;
				o = c.content.length;
				if (!o || o >= 4096) continue;
				c.start = s;
				f(o + 63 >> 6)
			}
			while (n.l & 511) n._W(-4, H.ENDOFCHAIN);
			for (i = 0; i < a[4] << 2; ++i) {
				var u = e.FullPaths[i];
				if (!u || u.length === 0) {
					for (l = 0; l < 17; ++l) n._W(4, 0);
					for (l = 0; l < 3; ++l) n._W(4, -1);
					for (l = 0; l < 12; ++l) n._W(4, 0);
					continue
				}
				c = e.FileIndex[i];
				if (i === 0) c.start = c.size ? c.start - 1 : L;
				var h = i === 0 && t.root || c.name;
				if (h.length > 32) {
					console.error("Name " + h + " will be truncated to " + h.slice(0, 32));
					h = h.slice(0, 32)
				}
				o = 2 * (h.length + 1);
				n._W(64, h, "utf16le");
				n._W(2, o);
				n._W(1, c.type);
				n._W(1, c.color);
				n._W(-4, c.L);
				n._W(-4, c.R);
				n._W(-4, c.C);
				if (!c.clsid) for (l = 0; l < 4; ++l) n._W(4, 0);
				else n._W(16, c.clsid, "hex");
				n._W(4, c.state || 0);
				n._W(4, 0);
				n._W(4, 0);
				n._W(4, 0);
				n._W(4, 0);
				n._W(4, c.start);
				n._W(4, c.size);
				n._W(4, 0)
			}
			for (i = 1; i < e.FileIndex.length; ++i) {
				c = e.FileIndex[i];
				if (c.size >= 4096) {
					n.l = c.start + 1 << 9;
					if (S && Buffer.isBuffer(c.content)) {
						c.content.copy(n, n.l, 0, c.size);
						n.l += c.size + 511 & -512
					} else {
						for (l = 0; l < c.size; ++l) n._W(1, c.content[l]);
						for (; l & 511; ++l) n._W(1, 0)
					}
				}
			}
			for (i = 1; i < e.FileIndex.length; ++i) {
				c = e.FileIndex[i];
				if (c.size > 0 && c.size < 4096) {
					if (S && Buffer.isBuffer(c.content)) {
						c.content.copy(n, n.l, 0, c.size);
						n.l += c.size + 63 & -64
					} else {
						for (l = 0; l < c.size; ++l) n._W(1, c.content[l]);
						for (; l & 63; ++l) n._W(1, 0)
					}
				}
			}
			if (S) {
				n.l = n.length
			} else {
				while (n.l < n.length) n._W(1, 0)
			}
			return n
		}
		function F(e, r) {
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
			var f = !i.match(U);
			i = i.replace(M, "");
			if (f) i = i.replace(U, "!");
			for (s = 0; s < t.length; ++s) {
				if ((f ? t[s].replace(U, "!") : t[s]).replace(M, "") == i) return e.FileIndex[s];
				if ((f ? a[s].replace(U, "!") : a[s]).replace(M, "") == i) return e.FileIndex[s]
			}
			return null
		}
		var D = 64;
		var L = -2;
		var B = "d0cf11e0a1b11ae1";
		var W = [208, 207, 17, 224, 161, 177, 26, 225];
		var z = "00000000000000000000000000000000";
		var H = {
			MAXREGSECT: -6,
			DIFSECT: -4,
			FATSECT: -3,
			ENDOFCHAIN: L,
			FREESECT: -1,
			HEADER_SIGNATURE: B,
			HEADER_MINOR_VERSION: "3e00",
			MAXREGSID: -6,
			NOSTREAM: -1,
			HEADER_CLSID: z,
			EntryTypes: ["unknown", "storage", "stream", "lockbytes", "property", "root"]
		};
		function V(e, r, t) {
			l();
			var a = N(e, t);
			f.writeFileSync(r, a)
		}
		function $(e) {
			var r = new Array(e.length);
			for (var t = 0; t < e.length; ++t) r[t] = String.fromCharCode(e[t]);
			return r.join("")
		}
		function X(e, r) {
			var t = N(e, r);
			switch (r && r.type || "buffer") {
			case "file":
				l();
				f.writeFileSync(r.filename, t);
				return t;
			case "binary":
				return typeof t == "string" ? t: $(t);
			case "base64":
				return T(typeof t == "string" ? t: $(t));
			case "buffer":
				if (S) return Buffer.isBuffer(t) ? t: x(t);
			case "array":
				return typeof t == "string" ? O(t) : t;
			}
			return t
		}
		var G;
		function j(e) {
			try {
				var r = e.InflateRaw;
				var t = new r;
				t._processChunk(new Uint8Array([3, 0]), t._finishFlushFlag);
				if (t.bytesRead) G = e;
				else throw new Error("zlib does not expose bytesRead")
			} catch(a) {
				console.error("cannot use native zlib: " + (a.message || a))
			}
		}
		function K(e, r) {
			if (!G) return Re(e, r);
			var t = G.InflateRaw;
			var a = new t;
			var n = a._processChunk(e.slice(e.l), a._finishFlushFlag);
			e.l += a.bytesRead;
			return n
		}
		function Y(e) {
			return G ? G.deflateRawSync(e) : Te(e)
		}
		var Z = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
		var J = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
		var q = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
		function Q(e) {
			var r = (e << 1 | e << 11) & 139536 | (e << 5 | e << 15) & 558144;
			return (r >> 16 | r >> 8 | r) & 255
		}
		var ee = typeof Uint8Array !== "undefined";
		var re = ee ? new Uint8Array(1 << 8) : [];
		for (var te = 0; te < 1 << 8; ++te) re[te] = Q(te);
		function ae(e, r) {
			var t = re[e & 255];
			if (r <= 8) return t >>> 8 - r;
			t = t << 8 | re[e >> 8 & 255];
			if (r <= 16) return t >>> 16 - r;
			t = t << 8 | re[e >> 16 & 255];
			return t >>> 24 - r
		}
		function ne(e, r) {
			var t = r & 7,
			a = r >>> 3;
			return (e[a] | (t <= 6 ? 0 : e[a + 1] << 8)) >>> t & 3
		}
		function ie(e, r) {
			var t = r & 7,
			a = r >>> 3;
			return (e[a] | (t <= 5 ? 0 : e[a + 1] << 8)) >>> t & 7
		}
		function se(e, r) {
			var t = r & 7,
			a = r >>> 3;
			return (e[a] | (t <= 4 ? 0 : e[a + 1] << 8)) >>> t & 15
		}
		function fe(e, r) {
			var t = r & 7,
			a = r >>> 3;
			return (e[a] | (t <= 3 ? 0 : e[a + 1] << 8)) >>> t & 31
		}
		function le(e, r) {
			var t = r & 7,
			a = r >>> 3;
			return (e[a] | (t <= 1 ? 0 : e[a + 1] << 8)) >>> t & 127
		}
		function oe(e, r, t) {
			var a = r & 7,
			n = r >>> 3,
			i = (1 << t) - 1;
			var s = e[n] >>> a;
			if (t < 8 - a) return s & i;
			s |= e[n + 1] << 8 - a;
			if (t < 16 - a) return s & i;
			s |= e[n + 2] << 16 - a;
			if (t < 24 - a) return s & i;
			s |= e[n + 3] << 24 - a;
			return s & i
		}
		function ce(e, r, t) {
			var a = r & 7,
			n = r >>> 3;
			if (a <= 5) e[n] |= (t & 7) << a;
			else {
				e[n] |= t << a & 255;
				e[n + 1] = (t & 7) >> 8 - a
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
		function ve(e, r) {
			var t = e.length,
			a = 2 * t > r ? 2 * t: r + 5,
			n = 0;
			if (t >= r) return e;
			if (S) {
				var i = R(a);
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
		function pe(e) {
			var r = new Array(e);
			for (var t = 0; t < e; ++t) r[t] = 0;
			return r
		}
		function me(e, r, t) {
			var a = 1,
			n = 0,
			i = 0,
			s = 0,
			f = 0,
			l = e.length;
			var o = ee ? new Uint16Array(32) : pe(32);
			for (i = 0; i < 32; ++i) o[i] = 0;
			for (i = l; i < t; ++i) e[i] = 0;
			l = e.length;
			var c = ee ? new Uint16Array(l) : pe(l);
			for (i = 0; i < l; ++i) {
				o[n = e[i]]++;
				if (a < n) a = n;
				c[i] = 0
			}
			o[0] = 0;
			for (i = 1; i <= a; ++i) o[i + 16] = f = f + o[i - 1] << 1;
			for (i = 0; i < l; ++i) {
				f = e[i];
				if (f != 0) c[i] = o[f + 16]++
			}
			var u = 0;
			for (i = 0; i < l; ++i) {
				u = e[i];
				if (u != 0) {
					f = ae(c[i], a) >> a - u;
					for (s = (1 << a + 4 - u) - 1; s >= 0; --s) r[f | s << u] = u & 15 | i << 4
				}
			}
			return a
		}
		var be = ee ? new Uint16Array(512) : pe(512);
		var ge = ee ? new Uint16Array(32) : pe(32);
		if (!ee) {
			for (var we = 0; we < 512; ++we) be[we] = 0;
			for (we = 0; we < 32; ++we) ge[we] = 0
		} (function() {
			var e = [];
			var r = 0;
			for (; r < 32; r++) e.push(5);
			me(e, ge, 32);
			var t = [];
			r = 0;
			for (; r <= 143; r++) t.push(8);
			for (; r <= 255; r++) t.push(9);
			for (; r <= 279; r++) t.push(7);
			for (; r <= 287; r++) t.push(8);
			me(t, be, 288)
		})();
		var ke = function je() {
			var e = ee ? new Uint8Array(32768) : [];
			var r = 0,
			t = 0;
			for (; r < q.length - 1; ++r) {
				for (; t < q[r + 1]; ++t) e[t] = r
			}
			for (; t < 32768; ++t) e[t] = 29;
			var a = ee ? new Uint8Array(259) : [];
			for (r = 0, t = 0; r < J.length - 1; ++r) {
				for (; t < J[r + 1]; ++t) a[t] = r
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
					var f = Math.min(65535, r.length - i);
					if (f < 10) {
						n = ce(t, n, +!!(i + f == r.length));
						if (n & 7) n += 8 - (n & 7);
						t.l = n / 8 | 0;
						t._W(2, f);
						t._W(2, ~f & 65535);
						while (f-->0) t[t.l++] = r[i++];
						n = t.l * 8;
						continue
					}
					n = ce(t, n, +!!(i + f == r.length) + 2);
					var l = 0;
					while (f-->0) {
						var o = r[i];
						l = (l << 5 ^ o) & 32767;
						var c = -1,
						u = 0;
						if (c = s[l]) {
							c |= i & ~32767;
							if (c > i) c -= 32768;
							if (c < i) while (r[c + u] == r[i + u] && u < 250)++u
						}
						if (u > 2) {
							o = a[u];
							if (o <= 22) n = he(t, n, re[o + 1] >> 1) - 1;
							else {
								he(t, n, 3);
								n += 5;
								he(t, n, re[o - 23] >> 5);
								n += 3
							}
							var h = o < 8 ? 0 : o - 4 >> 2;
							if (h > 0) {
								de(t, n, u - J[o]);
								n += h
							}
							o = e[i - c];
							n = he(t, n, re[o] >> 3);
							n -= 3;
							var d = o < 4 ? 0 : o - 2 >> 1;
							if (d > 0) {
								de(t, n, i - c - q[o]);
								n += d
							}
							for (var v = 0; v < u; ++v) {
								s[l] = i & 32767;
								l = (l << 5 ^ r[i]) & 32767; ++i
							}
							f -= u - 1
						} else {
							if (o <= 143) o = o + 48;
							else n = ue(t, n, 1);
							n = he(t, n, re[o]);
							s[l] = i & 32767; ++i
						}
					}
					n = he(t, n, 0) - 1
				}
				t.l = (n + 7) / 8 | 0;
				return t.l
			}
			return function s(e, r) {
				if (e.length < 8) return n(e, r);
				return i(e, r)
			}
		} ();
		function Te(e) {
			var r = Ea(50 + Math.floor(e.length * 1.1));
			var t = ke(e, r);
			return r.slice(0, t)
		}
		var ye = ee ? new Uint16Array(32768) : pe(32768);
		var Ee = ee ? new Uint16Array(32768) : pe(32768);
		var _e = ee ? new Uint16Array(128) : pe(128);
		var Se = 1,
		xe = 1;
		function Ae(e, r) {
			var t = fe(e, r) + 257;
			r += 5;
			var a = fe(e, r) + 1;
			r += 5;
			var n = se(e, r) + 4;
			r += 4;
			var i = 0;
			var s = ee ? new Uint8Array(19) : pe(19);
			var f = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
			var l = 1;
			var o = ee ? new Uint8Array(8) : pe(8);
			var c = ee ? new Uint8Array(8) : pe(8);
			var u = s.length;
			for (var h = 0; h < n; ++h) {
				s[Z[h]] = i = ie(e, r);
				if (l < i) l = i;
				o[i]++;
				r += 3
			}
			var d = 0;
			o[0] = 0;
			for (h = 1; h <= l; ++h) c[h] = d = d + o[h - 1] << 1;
			for (h = 0; h < u; ++h) if ((d = s[h]) != 0) f[h] = c[d]++;
			var v = 0;
			for (h = 0; h < u; ++h) {
				v = s[h];
				if (v != 0) {
					d = re[f[h]] >> 8 - v;
					for (var p = (1 << 7 - v) - 1; p >= 0; --p) _e[d | p << v] = v & 7 | h << 3
				}
			}
			var m = [];
			l = 1;
			for (; m.length < t + a;) {
				d = _e[le(e, r)];
				r += d & 7;
				switch (d >>>= 3) {
				case 16:
					i = 3 + ne(e, r);
					r += 2;
					d = m[m.length - 1];
					while (i-->0) m.push(d);
					break;
				case 17:
					i = 3 + ie(e, r);
					r += 3;
					while (i-->0) m.push(0);
					break;
				case 18:
					i = 11 + le(e, r);
					r += 7;
					while (i-->0) m.push(0);
					break;
				default:
					m.push(d);
					if (l < d) l = d;
					break;
				}
			}
			var b = m.slice(0, t),
			g = m.slice(t);
			for (h = t; h < 286; ++h) b[h] = 0;
			for (h = a; h < 30; ++h) g[h] = 0;
			Se = me(b, ye, 286);
			xe = me(g, Ee, 30);
			return r
		}
		function Ce(e, r) {
			if (e[0] == 3 && !(e[1] & 3)) {
				return [C(r), 2]
			}
			var t = 0;
			var a = 0;
			var n = R(r ? r: 1 << 18);
			var i = 0;
			var s = n.length >>> 0;
			var f = 0,
			l = 0;
			while ((a & 1) == 0) {
				a = ie(e, t);
				t += 3;
				if (a >>> 1 == 0) {
					if (t & 7) t += 8 - (t & 7);
					var o = e[t >>> 3] | e[(t >>> 3) + 1] << 8;
					t += 32;
					if (o > 0) {
						if (!r && s < i + o) {
							n = ve(n, i + o);
							s = n.length
						}
						while (o-->0) {
							n[i++] = e[t >>> 3];
							t += 8
						}
					}
					continue
				} else if (a >> 1 == 1) {
					f = 9;
					l = 5
				} else {
					t = Ae(e, t);
					f = Se;
					l = xe
				}
				for (;;) {
					if (!r && s < i + 32767) {
						n = ve(n, i + 32767);
						s = n.length
					}
					var c = oe(e, t, f);
					var u = a >>> 1 == 1 ? be[c] : ye[c];
					t += u & 15;
					u >>>= 4;
					if ((u >>> 8 & 255) === 0) n[i++] = u;
					else if (u == 256) break;
					else {
						u -= 257;
						var h = u < 8 ? 0 : u - 4 >> 2;
						if (h > 5) h = 0;
						var d = i + J[u];
						if (h > 0) {
							d += oe(e, t, h);
							t += h
						}
						c = oe(e, t, l);
						u = a >>> 1 == 1 ? ge[c] : Ee[c];
						t += u & 15;
						u >>>= 4;
						var v = u < 4 ? 0 : u - 2 >> 1;
						var p = q[u];
						if (v > 0) {
							p += oe(e, t, v);
							t += v
						}
						if (!r && s < d) {
							n = ve(n, d + 100);
							s = n.length
						}
						while (i < d) {
							n[i] = n[i - p]; ++i
						}
					}
				}
			}
			if (r) return [n, t + 7 >>> 3];
			return [n.slice(0, i), t + 7 >>> 3]
		}
		function Re(e, r) {
			var t = e.slice(e.l || 0);
			var a = Ce(t, r);
			e.l += a[1];
			return a[0]
		}
		function Oe(e, r) {
			if (e) {
				if (typeof console !== "undefined") console.error(r)
			} else throw new Error(r)
		}
		function Ie(e, r) {
			var t = e;
			Ta(t, 0);
			var a = [],
			n = [];
			var i = {
				FileIndex: a,
				FullPaths: n
			};
			E(i, {
				root: r.root
			});
			var f = t.length - 4;
			while ((t[f] != 80 || t[f + 1] != 75 || t[f + 2] != 5 || t[f + 3] != 6) && f >= 0)--f;
			t.l = f + 4;
			t.l += 4;
			var l = t._R(2);
			t.l += 6;
			var o = t._R(4);
			t.l = o;
			for (f = 0; f < l; ++f) {
				t.l += 20;
				var c = t._R(4);
				var u = t._R(4);
				var h = t._R(2);
				var d = t._R(2);
				var v = t._R(2);
				t.l += 8;
				var p = t._R(4);
				var m = s(t.slice(t.l + h, t.l + h + d));
				t.l += h + d + v;
				var b = t.l;
				t.l = p + 4;
				if (m && m[1]) {
					if ((m[1] || {}).usz) u = m[1].usz;
					if ((m[1] || {}).csz) c = m[1].csz
				}
				Ne(t, c, u, i, m);
				t.l = b
			}
			return i
		}
		function Ne(e, r, t, a, n) {
			e.l += 2;
			var f = e._R(2);
			var l = e._R(2);
			var o = i(e);
			if (f & 8257) throw new Error("Unsupported ZIP encryption");
			var c = e._R(4);
			var u = e._R(4);
			var h = e._R(4);
			var d = e._R(2);
			var v = e._R(2);
			var p = "";
			for (var m = 0; m < d; ++m) p += String.fromCharCode(e[e.l++]);
			if (v) {
				var b = s(e.slice(e.l, e.l + v));
				if ((b[21589] || {}).mt) o = b[21589].mt;
				if ((b[1] || {}).usz) h = b[1].usz;
				if ((b[1] || {}).csz) u = b[1].csz;
				if (n) {
					if ((n[21589] || {}).mt) o = n[21589].mt;
					if ((n[1] || {}).usz) h = b[1].usz;
					if ((n[1] || {}).csz) u = b[1].csz
				}
			}
			e.l += v;
			var g = e.slice(e.l, e.l + u);
			switch (l) {
			case 8:
				g = K(e, h);
				break;
			case 0:
				break;
			default:
				throw new Error("Unsupported ZIP Compression method " + l);
			}
			var w = false;
			if (f & 8) {
				c = e._R(4);
				if (c == 134695760) {
					c = e._R(4);
					w = true
				}
				u = e._R(4);
				h = e._R(4)
			}
			if (u != r) Oe(w, "Bad compressed size: " + r + " != " + u);
			if (h != t) Oe(w, "Bad uncompressed size: " + t + " != " + h);
			Ve(a, p, g, {
				unsafe: true,
				mt: o
			})
		}
		function Fe(e, r) {
			var t = r || {};
			var a = [],
			i = [];
			var s = Ea(1);
			var f = t.compression ? 8 : 0,
			l = 0;
			var o = false;
			if (o) l |= 8;
			var c = 0,
			u = 0;
			var h = 0,
			d = 0;
			var v = e.FullPaths[0],
			p = v,
			m = e.FileIndex[0];
			var b = [];
			var g = 0;
			for (c = 1; c < e.FullPaths.length; ++c) {
				p = e.FullPaths[c].slice(v.length);
				m = e.FileIndex[c];
				if (!m.size || !m.content || p == "Sh33tJ5") continue;
				var w = h;
				var k = Ea(p.length);
				for (u = 0; u < p.length; ++u) k._W(1, p.charCodeAt(u) & 127);
				k = k.slice(0, k.l);
				b[d] = typeof m.content == "string" ? qe.bstr(m.content, 0) : qe.buf(m.content, 0);
				var T = typeof m.content == "string" ? O(m.content) : m.content;
				if (f == 8) T = Y(T);
				s = Ea(30);
				s._W(4, 67324752);
				s._W(2, 20);
				s._W(2, l);
				s._W(2, f);
				if (m.mt) n(s, m.mt);
				else s._W(4, 0);
				s._W(-4, l & 8 ? 0 : b[d]);
				s._W(4, l & 8 ? 0 : T.length);
				s._W(4, l & 8 ? 0 : m.content.length);
				s._W(2, k.length);
				s._W(2, 0);
				h += s.length;
				a.push(s);
				h += k.length;
				a.push(k);
				h += T.length;
				a.push(T);
				if (l & 8) {
					s = Ea(12);
					s._W(-4, b[d]);
					s._W(4, T.length);
					s._W(4, m.content.length);
					h += s.l;
					a.push(s)
				}
				s = Ea(46);
				s._W(4, 33639248);
				s._W(2, 0);
				s._W(2, 20);
				s._W(2, l);
				s._W(2, f);
				s._W(4, 0);
				s._W(-4, b[d]);
				s._W(4, T.length);
				s._W(4, m.content.length);
				s._W(2, k.length);
				s._W(2, 0);
				s._W(2, 0);
				s._W(2, 0);
				s._W(2, 0);
				s._W(4, 0);
				s._W(4, w);
				g += s.l;
				i.push(s);
				g += k.length;
				i.push(k); ++d
			}
			s = Ea(22);
			s._W(4, 101010256);
			s._W(2, 0);
			s._W(2, 0);
			s._W(2, d);
			s._W(2, d);
			s._W(4, g);
			s._W(4, h);
			s._W(2, 0);
			return P([P(a), P(i), s])
		}
		var De = {
			htm: "text/html",
			xml: "text/xml",
			gif: "image/gif",
			jpg: "image/jpeg",
			png: "image/png",
			mso: "application/x-mso",
			thmx: "application/vnd.ms-officetheme",
			sh33tj5: "application/octet-stream"
		};
		function Pe(e, r) {
			if (e.ctype) return e.ctype;
			var t = e.name || "",
			a = t.match(/\.([^\.]+)$/);
			if (a && De[a[1]]) return De[a[1]];
			if (r) {
				a = (t = r).match(/[\.\\]([^\.\\])+$/);
				if (a && De[a[1]]) return De[a[1]]
			}
			return "application/octet-stream"
		}
		function Le(e) {
			var r = T(e);
			var t = [];
			for (var a = 0; a < r.length; a += 76) t.push(r.slice(a, a + 76));
			return t.join("\r\n") + "\r\n"
		}
		function Me(e) {
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
					var f = 76;
					var l = i.slice(s, s + f);
					if (l.charAt(f - 1) == "=") f--;
					else if (l.charAt(f - 2) == "=") f -= 2;
					else if (l.charAt(f - 3) == "=") f -= 3;
					l = i.slice(s, s + f);
					s += f;
					if (s < i.length) l += "=";
					t.push(l)
				}
			}
			return t.join("\r\n")
		}
		function Ue(e) {
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
			return O(r.join("\r\n"))
		}
		function Be(e, r, t) {
			var a = "",
			n = "",
			i = "",
			s;
			var f = 0;
			for (; f < 10; ++f) {
				var l = r[f];
				if (!l || l.match(/^\s*$/)) break;
				var o = l.match(/^(.*?):\s*([^\s].*)$/);
				if (o) switch (o[1].toLowerCase()) {
				case "content-location":
					a = o[2].trim();
					break;
				case "content-type":
					i = o[2].trim();
					break;
				case "content-transfer-encoding":
					n = o[2].trim();
					break;
				}
			}++f;
			switch (n.toLowerCase()) {
			case "base64":
				s = O(_(r.slice(f).join("")));
				break;
			case "quoted-printable":
				s = Ue(r.slice(f));
				break;
			default:
				throw new Error("Unsupported Content-Transfer-Encoding " + n);
			}
			var c = Ve(e, a.slice(t.length), s, {
				unsafe: true
			});
			if (i) c.ctype = i
		}
		function We(e, r) {
			if ($(e.slice(0, 13)).toLowerCase() != "mime-version:") throw new Error("Unsupported MAD header");
			var t = r && r.root || "";
			var a = (S && Buffer.isBuffer(e) ? e.toString("binary") : $(e)).split("\r\n");
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
			var f = "--" + (s[1] || "");
			var l = [],
			o = [];
			var c = {
				FileIndex: l,
				FullPaths: o
			};
			E(c);
			var u, h = 0;
			for (n = 0; n < a.length; ++n) {
				var d = a[n];
				if (d !== f && d !== f + "--") continue;
				if (h++) Be(c, a.slice(u, n), t);
				u = n
			}
			return c
		}
		function ze(e, r) {
			var t = r || {};
			var a = t.boundary || "SheetJS";
			a = "------=" + a;
			var n = ["MIME-Version: 1.0", 'Content-Type: multipart/related; boundary="' + a.slice(2) + '"', "", "", ""];
			var i = e.FullPaths[0],
			s = i,
			f = e.FileIndex[0];
			for (var l = 1; l < e.FullPaths.length; ++l) {
				s = e.FullPaths[l].slice(i.length);
				f = e.FileIndex[l];
				if (!f.size || !f.content || s == "Sh33tJ5") continue;
				s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7E-\xFF]/g,
				function(e) {
					return "_x" + e.charCodeAt(0).toString(16) + "_"
				}).replace(/[\u0080-\uFFFF]/g,
				function(e) {
					return "_u" + e.charCodeAt(0).toString(16) + "_"
				});
				var o = f.content;
				var c = S && Buffer.isBuffer(o) ? o.toString("binary") : $(o);
				var u = 0,
				h = Math.min(1024, c.length),
				d = 0;
				for (var v = 0; v <= h; ++v) if ((d = c.charCodeAt(v)) >= 32 && d < 128)++u;
				var p = u >= h * 4 / 5;
				n.push(a);
				n.push("Content-Location: " + (t.root || "file:///C:/SheetJS/") + s);
				n.push("Content-Transfer-Encoding: " + (p ? "quoted-printable": "base64"));
				n.push("Content-Type: " + Pe(f, s));
				n.push("");
				n.push(p ? Me(c) : Le(c))
			}
			n.push(a + "--\r\n");
			return n.join("\r\n")
		}
		function He(e) {
			var r = {};
			E(r, e);
			return r
		}
		function Ve(e, r, t, n) {
			var i = n && n.unsafe;
			if (!i) E(e);
			var s = !i && Qe.find(e, r);
			if (!s) {
				var f = e.FullPaths[0];
				if (r.slice(0, f.length) == f) f = r;
				else {
					if (f.slice(-1) != "/") f += "/";
					f = (f + r).replace("//", "/")
				}
				s = {
					name: a(r),
					type: 2
				};
				e.FileIndex.push(s);
				e.FullPaths.push(f);
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
		function $e(e, r) {
			E(e);
			var t = Qe.find(e, r);
			if (t) for (var a = 0; a < e.FileIndex.length; ++a) if (e.FileIndex[a] == t) {
				e.FileIndex.splice(a, 1);
				e.FullPaths.splice(a, 1);
				return true
			}
			return false
		}
		function Xe(e, r, t) {
			E(e);
			var n = Qe.find(e, r);
			if (n) for (var i = 0; i < e.FileIndex.length; ++i) if (e.FileIndex[i] == n) {
				e.FileIndex[i].name = a(t);
				e.FullPaths[i] = t;
				return true
			}
			return false
		}
		function Ge(e) {
			I(e, true)
		}
		e.find = F;
		e.read = y;
		e.parse = o;
		e.write = X;
		e.writeFile = V;
		e.utils = {
			cfb_new: He,
			cfb_add: Ve,
			cfb_del: $e,
			cfb_mov: Xe,
			cfb_gc: Ge,
			ReadShift: pa,
			CheckField: ka,
			prep_blob: Ta,
			bconcat: P,
			use_zlib: j,
			_deflateRaw: Te,
			_inflateRaw: Re,
			consts: H
		};
		return e
	} ();
	var er;
	function rr(e) {
		er = e
	}
	function tr(e) {
		if (typeof e === "string") return I(e);
		if (Array.isArray(e)) return F(e);
		return e
	}
	function ar(e, r, t) {
		if (typeof er !== "undefined" && er.writeFileSync) return t ? er.writeFileSync(e, r, t) : er.writeFileSync(e, r);
		if (typeof Deno !== "undefined") {
			if (t && typeof r == "string") switch (t) {
			case "utf8":
				r = new TextEncoder(t).encode(r);
				break;
			case "binary":
				r = I(r);
				break;
			default:
				throw new Error("Unsupported encoding " + t);
			}
			return Deno.writeFileSync(e, r)
		}
		var a = t == "utf8" ? Tt(r) : r;
		if (typeof IE_SaveFile !== "undefined") return IE_SaveFile(a, e);
		if (typeof Blob !== "undefined") {
			var n = new Blob([tr(a)], {
				type: "application/octet-stream"
			});
			if (typeof navigator !== "undefined" && navigator.msSaveBlob) return navigator.msSaveBlob(n, e);
			if (typeof saveAs !== "undefined") return saveAs(n, e);
			if (typeof URL !== "undefined" && typeof document !== "undefined" && document.createElement && URL.createObjectURL) {
				var i = URL.createObjectURL(n);
				if (typeof chrome === "object" && typeof(chrome.downloads || {}).download == "function") {
					if (URL.revokeObjectURL && typeof setTimeout !== "undefined") setTimeout(function() {
						URL.revokeObjectURL(i)
					},
					6e4);
					return chrome.downloads.download({
						url: i,
						filename: e,
						saveAs: true
					})
				}
				var s = document.createElement("a");
				if (s.download != null) {
					s.download = e;
					s.href = i;
					document.body.appendChild(s);
					s.click();
					document.body.removeChild(s);
					if (URL.revokeObjectURL && typeof setTimeout !== "undefined") setTimeout(function() {
						URL.revokeObjectURL(i)
					},
					6e4);
					return i
				}
			} else if (typeof URL !== "undefined" && !URL.createObjectURL && typeof chrome === "object") {
				var f = "data:application/octet-stream;base64," + E(new Uint8Array(tr(a)));
				return chrome.downloads.download({
					url: f,
					filename: e,
					saveAs: true
				})
			}
		}
		if (typeof $ !== "undefined" && typeof File !== "undefined" && typeof Folder !== "undefined") try {
			var l = File(e);
			l.open("w");
			l.encoding = "binary";
			if (Array.isArray(r)) r = N(r);
			l.write(r);
			l.close();
			return r
		} catch(o) {
			if (!o.message || !o.message.match(/onstruct/)) throw o
		}
		throw new Error("cannot save file " + e)
	}
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
			if (!a.message || !a.message.match(/onstruct/)) throw a
		}
		throw new Error("Cannot access file " + e)
	}
	function ir(e) {
		var r = Object.keys(e),
		t = [];
		for (var a = 0; a < r.length; ++a) if (Object.prototype.hasOwnProperty.call(e, r[a])) t.push(r[a]);
		return t
	}
	function sr(e, r) {
		var t = [],
		a = ir(e);
		for (var n = 0; n !== a.length; ++n) if (t[e[a[n]][r]] == null) t[e[a[n]][r]] = a[n];
		return t
	}
	function fr(e) {
		var r = [],
		t = ir(e);
		for (var a = 0; a !== t.length; ++a) r[e[t[a]]] = t[a];
		return r
	}
	function lr(e) {
		var r = [],
		t = ir(e);
		for (var a = 0; a !== t.length; ++a) r[e[t[a]]] = parseInt(t[a], 10);
		return r
	}
	function or(e) {
		var r = [],
		t = ir(e);
		for (var a = 0; a !== t.length; ++a) {
			if (r[e[t[a]]] == null) r[e[t[a]]] = [];
			r[e[t[a]]].push(t[a])
		}
		return r
	}
	var cr = Date.UTC(1899, 11, 30, 0, 0, 0);
	var ur = Date.UTC(1899, 11, 31, 0, 0, 0);
	var hr = Date.UTC(1904, 0, 1, 0, 0, 0);
	function dr(e, r) {
		var t = e.getTime();
		var a = (t - cr) / (24 * 60 * 60 * 1e3);
		if (r) {
			a -= 1462;
			return a < -1402 ? a - 1 : a
		}
		return a < 60 ? a - 1 : a
	}
	function vr(e) {
		if (e >= 60 && e < 61) return e;
		var r = new Date;
		r.setTime((e > 60 ? e: e + 1) * 24 * 60 * 60 * 1e3 + cr);
		return r
	}
	function pr(e) {
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
				break;
			}
			r += t * parseInt(n[i], 10)
		}
		return r
	}
	var mr = /^(\d+):(\d+)(:\d+)?(\.\d+)?$/;
	var br = /^(\d+)-(\d+)-(\d+)$/;
	var gr = /^(\d+)-(\d+)-(\d+)[T ](\d+):(\d+)(:\d+)?(\.\d+)?$/;
	function wr(e, r) {
		if (e instanceof Date) return e;
		var t = e.match(mr);
		if (t) return new Date((r ? hr: ur) + ((parseInt(t[1], 10) * 60 + parseInt(t[2], 10)) * 60 + (t[3] ? parseInt(t[3].slice(1), 10) : 0)) * 1e3 + (t[4] ? parseInt((t[4] + "000").slice(1, 4), 10) : 0));
		t = e.match(br);
		if (t) return new Date(Date.UTC( + t[1], +t[2] - 1, +t[3], 0, 0, 0, 0));
		t = e.match(gr);
		if (t) return new Date(Date.UTC( + t[1], +t[2] - 1, +t[3], +t[4], +t[5], t[6] && parseInt(t[6].slice(1), 10) || 0, t[7] && parseInt(t[7].slice(1), 10) || 0));
		var a = new Date(e);
		return a
	}
	function kr(e, r) {
		if (S && Buffer.isBuffer(e)) {
			if (r && A) {
				if (e[0] == 255 && e[1] == 254) return Tt(e.slice(2).toString("utf16le"));
				if (e[1] == 254 && e[2] == 255) return Tt(d(e.slice(2).toString("binary")))
			}
			return e.toString("binary")
		}
		if (typeof TextDecoder !== "undefined") try {
			if (r) {
				if (e[0] == 255 && e[1] == 254) return Tt(new TextDecoder("utf-16le").decode(e.slice(2)));
				if (e[0] == 254 && e[1] == 255) return Tt(new TextDecoder("utf-16be").decode(e.slice(2)))
			}
			var t = {
				"€": "",
				"‚": "",
				"ƒ": "",
				"„": "",
				"…": "",
				"†": "",
				"‡": "",
				"ˆ": "",
				"‰": "",
				"Š": "",
				"‹": "",
				"Œ": "",
				"Ž": "",
				"‘": "",
				"’": "",
				"“": "",
				"”": "",
				"•": "",
				"–": "",
				"—": "",
				"˜": "",
				"™": "",
				"š": "",
				"›": "",
				"œ": "",
				"ž": "",
				"Ÿ": ""
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
	function Tr(e) {
		if (typeof JSON != "undefined" && !Array.isArray(e)) return JSON.parse(JSON.stringify(e));
		if (typeof e != "object" || e == null) return e;
		if (e instanceof Date) return new Date(e.getTime());
		var r = {};
		for (var t in e) if (Object.prototype.hasOwnProperty.call(e, t)) r[t] = Tr(e[t]);
		return r
	}
	function yr(e, r) {
		var t = "";
		while (t.length < r) t += e;
		return t
	}
	function Er(e) {
		var r = Number(e);
		if (!isNaN(r)) return isFinite(r) ? r: NaN;
		if (!/\d/.test(e)) return r;
		var t = 1;
		var a = e.replace(/([\d]),([\d])/g, "$1$2").replace(/[$]/g, "").replace(/[%]/g,
		function() {
			t *= 100;
			return ""
		});
		if (!isNaN(r = Number(a))) return r / t;
		a = a.replace(/[(](.*)[)]/,
		function(e, r) {
			t = -t;
			return r
		});
		if (!isNaN(r = Number(a))) return r / t;
		return r
	}
	var _r = /^(0?\d|1[0-2])(?:|:([0-5]?\d)(?:|(\.\d+)(?:|:([0-5]?\d))|:([0-5]?\d)(|\.\d+)))\s+([ap])m?$/;
	var Sr = /^([01]?\d|2[0-3])(?:|:([0-5]?\d)(?:|(\.\d+)(?:|:([0-5]?\d))|:([0-5]?\d)(|\.\d+)))$/;
	var xr = /^(\d+)-(\d+)-(\d+)[T ](\d+):(\d+)(:\d+)(\.\d+)?[Z]?$/;
	var Ar = new Date("6/9/69 00:00 UTC").valueOf() == -177984e5;
	function Cr(e) {
		if (!e[2]) return new Date(Date.UTC(1899, 11, 31, +e[1] % 12 + (e[7] == "p" ? 12 : 0), 0, 0, 0));
		if (e[3]) {
			if (e[4]) return new Date(Date.UTC(1899, 11, 31, +e[1] % 12 + (e[7] == "p" ? 12 : 0), +e[2], +e[4], parseFloat(e[3]) * 1e3));
			else return new Date(Date.UTC(1899, 11, 31, e[7] == "p" ? 12 : 0, +e[1], +e[2], parseFloat(e[3]) * 1e3))
		} else if (e[5]) return new Date(Date.UTC(1899, 11, 31, +e[1] % 12 + (e[7] == "p" ? 12 : 0), +e[2], +e[5], e[6] ? parseFloat(e[6]) * 1e3: 0));
		else return new Date(Date.UTC(1899, 11, 31, +e[1] % 12 + (e[7] == "p" ? 12 : 0), +e[2], 0, 0))
	}
	function Rr(e) {
		if (!e[2]) return new Date(Date.UTC(1899, 11, 31, +e[1], 0, 0, 0));
		if (e[3]) {
			if (e[4]) return new Date(Date.UTC(1899, 11, 31, +e[1], +e[2], +e[4], parseFloat(e[3]) * 1e3));
			else return new Date(Date.UTC(1899, 11, 31, 0, +e[1], +e[2], parseFloat(e[3]) * 1e3))
		} else if (e[5]) return new Date(Date.UTC(1899, 11, 31, +e[1], +e[2], +e[5], e[6] ? parseFloat(e[6]) * 1e3: 0));
		else return new Date(Date.UTC(1899, 11, 31, +e[1], +e[2], 0, 0))
	}
	var Or = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
	function Ir(e) {
		if (xr.test(e)) return e.indexOf("Z") == -1 ? Dr(new Date(e)) : new Date(e);
		var r = e.toLowerCase();
		var t = r.replace(/\s+/g, " ").trim();
		var a = t.match(_r);
		if (a) return Cr(a);
		a = t.match(Sr);
		if (a) return Rr(a);
		a = t.match(gr);
		if (a) return new Date(Date.UTC( + a[1], +a[2] - 1, +a[3], +a[4], +a[5], a[6] && parseInt(a[6].slice(1), 10) || 0, a[7] && parseInt(a[7].slice(1), 10) || 0));
		var n = new Date(Ar && e.indexOf("UTC") == -1 ? e + " UTC": e),
		i = new Date(NaN);
		var s = n.getYear(),
		f = n.getMonth(),
		l = n.getDate();
		if (isNaN(l)) return i;
		if (r.match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)) {
			r = r.replace(/[^a-z]/g, "").replace(/([^a-z]|^)[ap]m?([^a-z]|$)/, "");
			if (r.length > 3 && Or.indexOf(r) == -1) return i
		} else if (r.replace(/[ap]m?/, "").match(/[a-z]/)) return i;
		if (s < 0 || s > 8099 || e.match(/[^-0-9:,\/\\\ ]/)) return i;
		return n
	}
	var Nr = function() {
		var e = "abacaba".split(/(:?b)/i).length == 5;
		return function r(t, a, n) {
			if (e || typeof a == "string") return t.split(a);
			var i = t.split(a),
			s = [i[0]];
			for (var f = 1; f < i.length; ++f) {
				s.push(n);
				s.push(i[f])
			}
			return s
		}
	} ();
	function Fr(e) {
		return new Date(e.getUTCFullYear(), e.getUTCMonth(), e.getUTCDate(), e.getUTCHours(), e.getUTCMinutes(), e.getUTCSeconds(), e.getUTCMilliseconds())
	}
	function Dr(e) {
		return new Date(Date.UTC(e.getFullYear(), e.getMonth(), e.getDate(), e.getHours(), e.getMinutes(), e.getSeconds(), e.getMilliseconds()))
	}
	function Pr(e) {
		if (!e) return null;
		if (e.content && e.type) return kr(e.content, true);
		if (e.data) return v(e.data);
		if (e.asNodeBuffer && S) return v(e.asNodeBuffer().toString("binary"));
		if (e.asBinary) return v(e.asBinary());
		if (e._data && e._data.getContent) return v(kr(Array.prototype.slice.call(e._data.getContent(), 0)));
		return null
	}
	function Lr(e) {
		if (!e) return null;
		if (e.data) return c(e.data);
		if (e.asNodeBuffer && S) return e.asNodeBuffer();
		if (e._data && e._data.getContent) {
			var r = e._data.getContent();
			if (typeof r == "string") return c(r);
			return Array.prototype.slice.call(r)
		}
		if (e.content && e.type) return e.content;
		return null
	}
	function Mr(e) {
		return e && e.name.slice(-4) === ".bin" ? Lr(e) : Pr(e)
	}
	function Ur(e, r) {
		var t = e.FullPaths || ir(e.files);
		var a = r.toLowerCase().replace(/[\/]/g, "\\"),
		n = a.replace(/\\/g, "/");
		for (var i = 0; i < t.length; ++i) {
			var s = t[i].replace(/^Root Entry[\/]/, "").toLowerCase();
			if (a == s || n == s) return e.files ? e.files[t[i]] : e.FileIndex[i]
		}
		return null
	}
	function Br(e, r) {
		var t = Ur(e, r);
		if (t == null) throw new Error("Cannot find file " + r + " in zip");
		return t
	}
	function Wr(e, r, t) {
		if (!t) return Mr(Br(e, r));
		if (!r) return null;
		try {
			return Wr(e, r)
		} catch(a) {
			return null
		}
	}
	function zr(e, r, t) {
		if (!t) return Pr(Br(e, r));
		if (!r) return null;
		try {
			return zr(e, r)
		} catch(a) {
			return null
		}
	}
	function Hr(e, r, t) {
		if (!t) return Lr(Br(e, r));
		if (!r) return null;
		try {
			return Hr(e, r)
		} catch(a) {
			return null
		}
	}
	function Vr(e) {
		var r = e.FullPaths || ir(e.files),
		t = [];
		for (var a = 0; a < r.length; ++a) if (r[a].slice(-1) != "/") t.push(r[a].replace(/^Root Entry[\/]/, ""));
		return t.sort()
	}
	function $r(e, r, t) {
		if (e.FullPaths) {
			if (typeof t == "string") {
				var a;
				if (S) a = x(t);
				else a = L(t);
				return Qe.utils.cfb_add(e, r, a)
			}
			Qe.utils.cfb_add(e, r, t)
		} else e.file(r, t)
	}
	function Xr() {
		return Qe.utils.cfb_new()
	}
	function Gr(e, r) {
		switch (r.type) {
		case "base64":
			return Qe.read(e, {
				type: "base64"
			});
		case "binary":
			return Qe.read(e, {
				type: "binary"
			});
		case "buffer":
			;
		case "array":
			return Qe.read(e, {
				type: "buffer"
			});
		}
		throw new Error("Unrecognized type " + r.type)
	}
	function jr(e, r) {
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
	var Kr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
	var Yr = /([^"\s?>\/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;
	var Zr = /<[\/\?]?[a-zA-Z0-9:_-]+(?:\s+[^"\s?>\/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^'">\s=]+))*\s*[\/\?]?>/gm,
	Jr = /<[^>]*>/g;
	var qr = Kr.match(Zr) ? Zr: Jr;
	var Qr = /<\w*:/,
	et = /<(\/?)\w+:/;
	function rt(e, r, t) {
		var a = {};
		var n = 0,
		i = 0;
		for (; n !== e.length; ++n) if ((i = e.charCodeAt(n)) === 32 || i === 10 || i === 13) break;
		if (!r) a[0] = e.slice(0, n);
		if (n === e.length) return a;
		var s = e.match(Yr),
		f = 0,
		l = "",
		o = 0,
		c = "",
		u = "",
		h = 1;
		if (s) for (o = 0; o != s.length; ++o) {
			u = s[o];
			for (i = 0; i != u.length; ++i) if (u.charCodeAt(i) === 61) break;
			c = u.slice(0, i).trim();
			while (u.charCodeAt(i + 1) == 32)++i;
			h = (n = u.charCodeAt(i + 1)) == 34 || n == 39 ? 1 : 0;
			l = u.slice(i + 1 + h, u.length - h);
			for (f = 0; f != c.length; ++f) if (c.charCodeAt(f) === 58) break;
			if (f === c.length) {
				if (c.indexOf("_") > 0) c = c.slice(0, c.indexOf("_"));
				a[c] = l;
				if (!t) a[c.toLowerCase()] = l
			} else {
				var d = (f === 5 && c.slice(0, 5) === "xmlns" ? "xmlns": "") + c.slice(f + 1);
				if (a[d] && c.slice(f - 3, f) == "ext") continue;
				a[d] = l;
				if (!t) a[d.toLowerCase()] = l
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
		"&amp;": "&"
	};
	var nt = fr(at);
	var it = function() {
		var e = /&(?:quot|apos|gt|lt|amp|#x?([\da-fA-F]+));/gi,
		r = /_x([\da-fA-F]{4})_/gi;
		function t(a) {
			var n = a + "",
			i = n.indexOf("<![CDATA[");
			if (i == -1) return n.replace(e,
			function(e, r) {
				return at[e] || String.fromCharCode(parseInt(r, e.indexOf("x") > -1 ? 16 : 10)) || e
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
	} ();
	var st = /[&<>'"]/g,
	ft = /[\u0000-\u0008\u000b-\u001f\uFFFE-\uFFFF]/g;
	function lt(e) {
		var r = e + "";
		return r.replace(st,
		function(e) {
			return nt[e]
		}).replace(ft,
		function(e) {
			return "_x" + ("000" + e.charCodeAt(0).toString(16)).slice(-4) + "_"
		})
	}
	function ot(e) {
		return lt(e).replace(/ /g, "_x0020_")
	}
	var ct = /[\u0000-\u001f]/g;
	function ut(e) {
		var r = e + "";
		return r.replace(st,
		function(e) {
			return nt[e]
		}).replace(/\n/g, "<br/>").replace(ct,
		function(e) {
			return "&#x" + ("000" + e.charCodeAt(0).toString(16)).slice(-4) + ";"
		})
	}
	function ht(e) {
		var r = e + "";
		return r.replace(st,
		function(e) {
			return nt[e]
		}).replace(ct,
		function(e) {
			return "&#x" + e.charCodeAt(0).toString(16).toUpperCase() + ";"
		})
	}
	var dt = function() {
		var e = /&#(\d+);/g;
		function r(e, r) {
			return String.fromCharCode(parseInt(r, 10))
		}
		return function t(a) {
			return a.replace(e, r)
		}
	} ();
	function vt(e) {
		return e.replace(/(\r\n|[\r\n])/g, "&#10;")
	}
	function pt(e) {
		switch (e) {
		case 1:
			;
		case true:
			;
		case "1":
			;
		case "true":
			return true;
		case 0:
			;
		case false:
			;
		case "0":
			;
		case "false":
			return false;
		}
		return false
	}
	function mt(e) {
		var r = "",
		t = 0,
		a = 0,
		n = 0,
		i = 0,
		s = 0,
		f = 0;
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
				r += String.fromCharCode((a & 15) << 12 | (n & 63) << 6 | i & 63);
				continue
			}
			s = e.charCodeAt(t++);
			f = ((a & 7) << 18 | (n & 63) << 12 | (i & 63) << 6 | s & 63) - 65536;
			r += String.fromCharCode(55296 + (f >>> 10 & 1023));
			r += String.fromCharCode(56320 + (f & 1023))
		}
		return r
	}
	function bt(e) {
		var r = C(2 * e.length),
		t,
		a,
		n = 1,
		i = 0,
		s = 0,
		f;
		for (a = 0; a < e.length; a += n) {
			n = 1;
			if ((f = e.charCodeAt(a)) < 128) t = f;
			else if (f < 224) {
				t = (f & 31) * 64 + (e.charCodeAt(a + 1) & 63);
				n = 2
			} else if (f < 240) {
				t = (f & 15) * 4096 + (e.charCodeAt(a + 1) & 63) * 64 + (e.charCodeAt(a + 2) & 63);
				n = 3
			} else {
				n = 4;
				t = (f & 7) * 262144 + (e.charCodeAt(a + 1) & 63) * 4096 + (e.charCodeAt(a + 2) & 63) * 64 + (e.charCodeAt(a + 3) & 63);
				t -= 65536;
				s = 55296 + (t >>> 10 & 1023);
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
	function gt(e) {
		return x(e, "binary").toString("utf8")
	}
	var wt = "foo bar bazâð£";
	var kt = S && (gt(wt) == mt(wt) && gt || bt(wt) == mt(wt) && bt) || mt;
	var Tt = S ?
	function(e) {
		return x(e, "utf8").toString("binary")
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
				r.push(String.fromCharCode(240 + (n >> 18 & 7)));
				r.push(String.fromCharCode(144 + (n >> 12 & 63)));
				r.push(String.fromCharCode(128 + (n >> 6 & 63)));
				r.push(String.fromCharCode(128 + (n & 63)));
				break;
			default:
				r.push(String.fromCharCode(224 + (a >> 12)));
				r.push(String.fromCharCode(128 + (a >> 6 & 63)));
				r.push(String.fromCharCode(128 + (a & 63)));
			}
		}
		return r.join("")
	};
	var yt = function() {
		var e = {};
		return function r(t, a) {
			var n = t + "|" + (a || "");
			if (e[n]) return e[n];
			return e[n] = new RegExp("<(?:\\w+:)?" + t + '(?: xml:space="preserve")?(?:[^>]*)>([\\s\\S]*?)</(?:\\w+:)?' + t + ">", a || "")
		}
	} ();
	var Et = function() {
		var e = [["nbsp", " "], ["middot", "·"], ["quot", '"'], ["apos", "'"], ["gt", ">"], ["lt", "<"], ["amp", "&"]].map(function(e) {
			return [new RegExp("&" + e[0] + ";", "ig"), e[1]]
		});
		return function r(t) {
			var a = t.replace(/^[\t\n\r ]+/, "").replace(/[\t\n\r ]+$/, "").replace(/>\s+/g, ">").replace(/\s+</g, "<").replace(/[\t\n\r ]+/g, " ").replace(/<\s*[bB][rR]\s*\/?>/g, "\n").replace(/<[^>]*>/g, "");
			for (var n = 0; n < e.length; ++n) a = a.replace(e[n][0], e[n][1]);
			return a
		}
	} ();
	var _t = function() {
		var e = {};
		return function r(t) {
			if (e[t] !== undefined) return e[t];
			return e[t] = new RegExp("<(?:vt:)?" + t + ">([\\s\\S]*?)</(?:vt:)?" + t + ">", "g")
		}
	} ();
	var St = /<\/?(?:vt:)?variant>/g,
	xt = /<(?:vt:)([^>]*)>([\s\S]*)</;
	function At(e, r) {
		var t = rt(e);
		var a = e.match(_t(t.baseType)) || [];
		var n = [];
		if (a.length != t.size) {
			if (r.WTF) throw new Error("unexpected vector length " + a.length + " != " + t.size);
			return n
		}
		a.forEach(function(e) {
			var r = e.replace(St, "").match(xt);
			if (r) n.push({
				v: kt(r[2]),
				t: r[1]
			})
		});
		return n
	}
	var Ct = /(^\s|\s$|\n)/;
	function Rt(e, r) {
		return "<" + e + (r.match(Ct) ? ' xml:space="preserve"': "") + ">" + r + "</" + e + ">"
	}
	function Ot(e) {
		return ir(e).map(function(r) {
			return " " + r + '="' + e[r] + '"'
		}).join("")
	}
	function It(e, r, t) {
		return "<" + e + (t != null ? Ot(t) : "") + (r != null ? (r.match(Ct) ? ' xml:space="preserve"': "") + ">" + r + "</" + e: "/") + ">"
	}
	function Nt(e, r) {
		try {
			return e.toISOString().replace(/\.\d*/, "")
		} catch(t) {
			if (r) throw t
		}
		return ""
	}
	function Ft(e, r) {
		switch (typeof e) {
		case "string":
			var t = It("vt:lpwstr", lt(e));
			if (r) t = t.replace(/&quot;/g, "_x0022_");
			return t;
		case "number":
			return It((e | 0) == e ? "vt:i4": "vt:r8", lt(String(e)));
		case "boolean":
			return It("vt:bool", e ? "true": "false");
		}
		if (e instanceof Date) return It("vt:filetime", Nt(e));
		throw new Error("Unable to serialize " + e)
	}
	function Dt(e) {
		if (S && Buffer.isBuffer(e)) return e.toString("utf8");
		if (typeof e === "string") return e;
		if (typeof Uint8Array !== "undefined" && e instanceof Uint8Array) return kt(N(D(e)));
		throw new Error("Bad input format: expected Buffer or string")
	}
	var Pt = /<(\/?)([^\s?><!\/:]*:|)([^\s?<>:\/]+)(?:[\s?:\/](?:[^>=]|="[^"]*?")*)?>/gm;
	var Lt = {
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
		xsd: "http://www.w3.org/2001/XMLSchema"
	};
	var Mt = ["http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main", "http://schemas.microsoft.com/office/excel/2006/main", "http://schemas.microsoft.com/office/excel/2006/2"];
	var Ut = {
		o: "urn:schemas-microsoft-com:office:office",
		x: "urn:schemas-microsoft-com:office:excel",
		ss: "urn:schemas-microsoft-com:office:spreadsheet",
		dt: "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882",
		mv: "http://macVmlSchemaUri",
		v: "urn:schemas-microsoft-com:vml",
		html: "http://www.w3.org/TR/REC-html40"
	};
	function Bt(e, r) {
		var t = 1 - 2 * (e[r + 7] >>> 7);
		var a = ((e[r + 7] & 127) << 4) + (e[r + 6] >>> 4 & 15);
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
	function Wt(e, r, t) {
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
		for (var f = 0; f <= 5; ++f, i /= 256) e[t + f] = i & 255;
		e[t + 6] = (n & 15) << 4 | i & 15;
		e[t + 7] = n >> 4 | a
	}
	var zt = function(e) {
		var r = [],
		t = 10240;
		for (var a = 0; a < e[0].length; ++a) if (e[0][a]) for (var n = 0,
		i = e[0][a].length; n < i; n += t) r.push.apply(r, e[0][a].slice(n, n + t));
		return r
	};
	var Ht = S ?
	function(e) {
		return e[0].length > 0 && Buffer.isBuffer(e[0][0]) ? Buffer.concat(e[0].map(function(e) {
			return Buffer.isBuffer(e) ? e: x(e)
		})) : zt(e)
	}: zt;
	var Vt = function(e, r, t) {
		var a = [];
		for (var n = r; n < t; n += 2) a.push(String.fromCharCode(ca(e, n)));
		return a.join("").replace(M, "")
	};
	var $t = S ?
	function(e, r, t) {
		if (!Buffer.isBuffer(e) || !A) return Vt(e, r, t);
		return e.toString("utf16le", r, t).replace(M, "")
	}: Vt;
	var Xt = function(e, r, t) {
		var a = [];
		for (var n = r; n < r + t; ++n) a.push(("0" + e[n].toString(16)).slice(-2));
		return a.join("")
	};
	var Gt = S ?
	function(e, r, t) {
		return Buffer.isBuffer(e) ? e.toString("hex", r, r + t) : Xt(e, r, t)
	}: Xt;
	var jt = function(e, r, t) {
		var a = [];
		for (var n = r; n < t; n++) a.push(String.fromCharCode(oa(e, n)));
		return a.join("")
	};
	var Kt = S ?
	function vT(e, r, t) {
		return Buffer.isBuffer(e) ? e.toString("utf8", r, t) : jt(e, r, t)
	}: jt;
	var Yt = function(e, r) {
		var t = ha(e, r);
		return t > 0 ? Kt(e, r + 4, r + 4 + t - 1) : ""
	};
	var Zt = Yt;
	var Jt = function(e, r) {
		var t = ha(e, r);
		return t > 0 ? Kt(e, r + 4, r + 4 + t - 1) : ""
	};
	var qt = Jt;
	var Qt = function(e, r) {
		var t = 2 * ha(e, r);
		return t > 0 ? Kt(e, r + 4, r + 4 + t - 1) : ""
	};
	var ea = Qt;
	var ra = function pT(e, r) {
		var t = ha(e, r);
		return t > 0 ? $t(e, r + 4, r + 4 + t) : ""
	};
	var ta = ra;
	var aa = function(e, r) {
		var t = ha(e, r);
		return t > 0 ? Kt(e, r + 4, r + 4 + t) : ""
	};
	var na = aa;
	var ia = function(e, r) {
		return Bt(e, r)
	};
	var sa = ia;
	var fa = function mT(e) {
		return Array.isArray(e) || typeof Uint8Array !== "undefined" && e instanceof Uint8Array
	};
	if (S) {
		Zt = function bT(e, r) {
			if (!Buffer.isBuffer(e)) return Yt(e, r);
			var t = e.readUInt32LE(r);
			return t > 0 ? e.toString("utf8", r + 4, r + 4 + t - 1) : ""
		};
		qt = function gT(e, r) {
			if (!Buffer.isBuffer(e)) return Jt(e, r);
			var t = e.readUInt32LE(r);
			return t > 0 ? e.toString("utf8", r + 4, r + 4 + t - 1) : ""
		};
		ea = function wT(e, r) {
			if (!Buffer.isBuffer(e) || !A) return Qt(e, r);
			var t = 2 * e.readUInt32LE(r);
			return e.toString("utf16le", r + 4, r + 4 + t - 1)
		};
		ta = function kT(e, r) {
			if (!Buffer.isBuffer(e) || !A) return ra(e, r);
			var t = e.readUInt32LE(r);
			return e.toString("utf16le", r + 4, r + 4 + t)
		};
		na = function TT(e, r) {
			if (!Buffer.isBuffer(e)) return aa(e, r);
			var t = e.readUInt32LE(r);
			return e.toString("utf8", r + 4, r + 4 + t)
		};
		sa = function yT(e, r) {
			if (Buffer.isBuffer(e)) return e.readDoubleLE(r);
			return ia(e, r)
		};
		fa = function ET(e) {
			return Buffer.isBuffer(e) || Array.isArray(e) || typeof Uint8Array !== "undefined" && e instanceof Uint8Array
		}
	}
	function la() {
		$t = function(e, r, t) {
			return a.utils.decode(1200, e.slice(r, t)).replace(M, "")
		};
		Kt = function(e, r, t) {
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
	if (typeof a !== "undefined") la();
	var oa = function(e, r) {
		return e[r]
	};
	var ca = function(e, r) {
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
		return e[r + 3] << 24 | e[r + 2] << 16 | e[r + 1] << 8 | e[r]
	};
	var va = function(e, r) {
		return e[r] << 24 | e[r + 1] << 16 | e[r + 2] << 8 | e[r + 3]
	};
	function pa(e, t) {
		var n = "",
		i, s, f = [],
		l,
		o,
		c,
		u;
		switch (t) {
		case "dbcs":
			u = this.l;
			if (S && Buffer.isBuffer(this) && A) n = this.slice(this.l, this.l + 2 * e).toString("utf16le");
			else for (c = 0; c < e; ++c) {
				n += String.fromCharCode(ca(this, u));
				u += 2
			}
			e *= 2;
			break;
		case "utf8":
			n = Kt(this, this.l, this.l + e);
			break;
		case "utf16le":
			e *= 2;
			n = $t(this, this.l, this.l + e);
			break;
		case "wstr":
			if (typeof a !== "undefined") n = a.utils.decode(r, this.slice(this.l, this.l + 2 * e));
			else return pa.call(this, e, "dbcs");
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
			while ((l = oa(this, this.l + e++)) !== 0) f.push(p(l));
			n = f.join("");
			break;
		case "_wstr":
			e = 0;
			n = "";
			while ((l = ca(this, this.l + e)) !== 0) {
				f.push(p(l));
				e += 2
			}
			e += 2;
			n = f.join("");
			break;
		case "dbcs-cont":
			n = "";
			u = this.l;
			for (c = 0; c < e; ++c) {
				if (this.lens && this.lens.indexOf(u) !== -1) {
					l = oa(this, u);
					this.l = u + 1;
					o = pa.call(this, e - c, l ? "dbcs-cont": "sbcs-cont");
					return f.join("") + o
				}
				f.push(p(ca(this, u)));
				u += 2
			}
			n = f.join("");
			e *= 2;
			break;
		case "cpstr":
			if (typeof a !== "undefined") {
				n = a.utils.decode(r, this.slice(this.l, this.l + e));
				break
			};
		case "sbcs-cont":
			n = "";
			u = this.l;
			for (c = 0; c != e; ++c) {
				if (this.lens && this.lens.indexOf(u) !== -1) {
					l = oa(this, u);
					this.l = u + 1;
					o = pa.call(this, e - c, l ? "dbcs-cont": "sbcs-cont");
					return f.join("") + o
				}
				f.push(p(oa(this, u)));
				u += 1
			}
			n = f.join("");
			break;
		default:
			switch (e) {
			case 1:
				i = oa(this, this.l);
				this.l++;
				return i;
			case 2:
				i = (t === "i" ? ua: ca)(this, this.l);
				this.l += 2;
				return i;
			case 4:
				;
			case - 4 : if (t === "i" || (this[this.l + 3] & 128) === 0) {
					i = (e > 0 ? da: va)(this, this.l);
					this.l += 4;
					return i
				} else {
					s = ha(this, this.l);
					this.l += 4
				}
				return s;
			case 8:
				;
			case - 8 : if (t === "f") {
					if (e == 8) s = sa(this, this.l);
					else s = sa([this[this.l + 7], this[this.l + 6], this[this.l + 5], this[this.l + 4], this[this.l + 3], this[this.l + 2], this[this.l + 1], this[this.l + 0]], 0);
					this.l += 8;
					return s
				} else e = 8;
			case 16:
				n = Gt(this, this.l, e);
				break;
			};
		}
		this.l += e;
		return n
	}
	var ma = function(e, r, t) {
		e[t] = r & 255;
		e[t + 1] = r >>> 8 & 255;
		e[t + 2] = r >>> 16 & 255;
		e[t + 3] = r >>> 24 & 255
	};
	var ba = function(e, r, t) {
		e[t] = r & 255;
		e[t + 1] = r >> 8 & 255;
		e[t + 2] = r >> 16 & 255;
		e[t + 3] = r >> 24 & 255
	};
	var ga = function(e, r, t) {
		e[t] = r & 255;
		e[t + 1] = r >>> 8 & 255
	};
	function wa(e, n, i) {
		var s = 0,
		f = 0;
		if (i === "dbcs") {
			for (f = 0; f != n.length; ++f) ga(this, n.charCodeAt(f), this.l + 2 * f);
			s = 2 * n.length
		} else if (i === "sbcs" || i == "cpstr") {
			if (typeof a !== "undefined" && t == 874) {
				for (f = 0; f != n.length; ++f) {
					var l = a.utils.encode(t, n.charAt(f));
					this[this.l + f] = l[0]
				}
				s = n.length
			} else if (typeof a !== "undefined" && i == "cpstr") {
				l = a.utils.encode(r, n);
				if (l.length == n.length) for (f = 0; f < n.length; ++f) if (l[f] == 0 && n.charCodeAt(f) != 0) l[f] = 95;
				if (l.length == 2 * n.length) for (f = 0; f < n.length; ++f) if (l[2 * f] == 0 && l[2 * f + 1] == 0 && n.charCodeAt(f) != 0) l[2 * f] = 95;
				for (f = 0; f < l.length; ++f) this[this.l + f] = l[f];
				s = l.length
			} else {
				n = n.replace(/[^\x00-\x7F]/g, "_");
				for (f = 0; f != n.length; ++f) this[this.l + f] = n.charCodeAt(f) & 255;
				s = n.length
			}
		} else if (i === "hex") {
			for (; f < e; ++f) {
				this[this.l++] = parseInt(n.slice(2 * f, 2 * f + 2), 16) || 0
			}
			return this
		} else if (i === "utf16le") {
			var o = Math.min(this.l + e, this.length);
			for (f = 0; f < Math.min(n.length, e); ++f) {
				var c = n.charCodeAt(f);
				this[this.l++] = c & 255;
				this[this.l++] = c >> 8
			}
			while (this.l < o) this[this.l++] = 0;
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
			ma(this, n, this.l);
			break;
		case 8:
			s = 8;
			if (i === "f") {
				Wt(this, n, this.l);
				break
			};
		case 16:
			break;
		case - 4 : s = 4;
			ba(this, n, this.l);
			break;
		}
		this.l += s;
		return this
	}
	function ka(e, r) {
		var t = Gt(this, this.l, e.length >> 1);
		if (t !== e) throw new Error(r + "Expected " + e + " saw " + t);
		this.l += e.length >> 1
	}
	function Ta(e, r) {
		e.l = r;
		e._R = pa;
		e.chk = ka;
		e._W = wa
	}
	function ya(e, r) {
		e.l += r
	}
	function Ea(e) {
		var r = C(e);
		Ta(r, 0);
		return r
	}
	function _a(e, r, t) {
		if (!e) return;
		var a, n, i;
		Ta(e, e.l || 0);
		var s = e.length,
		f = 0,
		l = 0;
		while (e.l < s) {
			f = e._R(1);
			if (f & 128) f = (f & 127) + ((e._R(1) & 127) << 7);
			var o = Qb[f] || Qb[65535];
			a = e._R(1);
			i = a & 127;
			for (n = 1; n < 4 && a & 128; ++n) i += ((a = e._R(1)) & 127) << 7 * n;
			l = e.l + i;
			var c = o.f && o.f(e, i, t);
			e.l = l;
			if (r(c, o, f)) return
		}
	}
	function Sa() {
		var e = [],
		r = S ? 256 : 2048;
		var t = function l(e) {
			var r = Ea(e);
			Ta(r, 0);
			return r
		};
		var a = t(r);
		var n = function o() {
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
		var i = function c(e) {
			if (a && e < a.length - a.l) return a;
			n();
			return a = t(Math.max(e + 1, r))
		};
		var s = function u() {
			n();
			return P(e)
		};
		var f = function h(e) {
			n();
			a = e;
			if (a.l == null) a.l = a.length;
			i(r)
		};
		return {
			next: i,
			push: f,
			end: s,
			_bufs: e
		}
	}
	function xa(e, r, t, a) {
		var n = +r,
		i;
		if (isNaN(n)) return;
		if (!a) a = Qb[n].p || (t || []).length || 0;
		i = 1 + (n >= 128 ? 1 : 0) + 1;
		if (a >= 128)++i;
		if (a >= 16384)++i;
		if (a >= 2097152)++i;
		var s = e.next(i);
		if (n <= 127) s._W(1, n);
		else {
			s._W(1, (n & 127) + 128);
			s._W(1, n >> 7)
		}
		for (var f = 0; f != 4; ++f) {
			if (a >= 128) {
				s._W(1, (a & 127) + 128);
				a >>= 7
			} else {
				s._W(1, a);
				break
			}
		}
		if (a > 0 && fa(t)) e.push(t)
	}
	function Aa(e, r, t) {
		var a = Tr(e);
		if (r.s) {
			if (a.cRel) a.c += r.s.c;
			if (a.rRel) a.r += r.s.r
		} else {
			if (a.cRel) a.c += r.c;
			if (a.rRel) a.r += r.r
		}
		if (!t || t.biff < 12) {
			while (a.c >= 256) a.c -= 256;
			while (a.r >= 65536) a.r -= 65536
		}
		return a
	}
	function Ca(e, r, t) {
		var a = Tr(e);
		a.s = Aa(a.s, r.s, t);
		a.e = Aa(a.e, r.s, t);
		return a
	}
	function Ra(e, r) {
		if (e.cRel && e.c < 0) {
			e = Tr(e);
			while (e.c < 0) e.c += r > 8 ? 16384 : 256
		}
		if (e.rRel && e.r < 0) {
			e = Tr(e);
			while (e.r < 0) e.r += r > 8 ? 1048576 : r > 5 ? 65536 : 16384
		}
		var t = za(e);
		if (!e.cRel && e.cRel != null) t = Ma(t);
		if (!e.rRel && e.rRel != null) t = Fa(t);
		return t
	}
	function Oa(e, r) {
		if (e.s.r == 0 && !e.s.rRel) {
			if (e.e.r == (r.biff >= 12 ? 1048575 : r.biff >= 8 ? 65536 : 16384) && !e.e.rRel) {
				return (e.s.cRel ? "": "$") + La(e.s.c) + ":" + (e.e.cRel ? "": "$") + La(e.e.c)
			}
		}
		if (e.s.c == 0 && !e.s.cRel) {
			if (e.e.c == (r.biff >= 12 ? 16383 : 255) && !e.e.cRel) {
				return (e.s.rRel ? "": "$") + Na(e.s.r) + ":" + (e.e.rRel ? "": "$") + Na(e.e.r)
			}
		}
		return Ra(e.s, r.biff) + ":" + Ra(e.e, r.biff)
	}
	if (typeof cptable !== "undefined") b(cptable);
	else if (typeof module !== "undefined" && typeof require !== "undefined") {
		b(undefined)
	}
	function Ia(e) {
		return parseInt(Da(e), 10) - 1
	}
	function Na(e) {
		return "" + (e + 1)
	}
	function Fa(e) {
		return e.replace(/([A-Z]|^)(\d+)$/, "$1$$$2")
	}
	function Da(e) {
		return e.replace(/\$(\d+)$/, "$1")
	}
	function Pa(e) {
		var r = Ua(e),
		t = 0,
		a = 0;
		for (; a !== r.length; ++a) t = 26 * t + r.charCodeAt(a) - 64;
		return t - 1
	}
	function La(e) {
		if (e < 0) throw new Error("invalid column " + e);
		var r = "";
		for (++e; e; e = Math.floor((e - 1) / 26)) r = String.fromCharCode((e - 1) % 26 + 65) + r;
		return r
	}
	function Ma(e) {
		return e.replace(/^([A-Z])/, "$$$1")
	}
	function Ua(e) {
		return e.replace(/^\$([A-Z])/, "$1")
	}
	function Ba(e) {
		return e.replace(/(\$?[A-Z]*)(\$?\d*)/, "$1,$2").split(",")
	}
	function Wa(e) {
		var r = 0,
		t = 0;
		for (var a = 0; a < e.length; ++a) {
			var n = e.charCodeAt(a);
			if (n >= 48 && n <= 57) r = 10 * r + (n - 48);
			else if (n >= 65 && n <= 90) t = 26 * t + (n - 64)
		}
		return {
			c: t - 1,
			r: r - 1
		}
	}
	function za(e) {
		var r = e.c + 1;
		var t = "";
		for (; r; r = (r - 1) / 26 | 0) t = String.fromCharCode((r - 1) % 26 + 65) + t;
		return t + (e.r + 1)
	}
	function Ha(e) {
		var r = e.indexOf(":");
		if (r == -1) return {
			s: Wa(e),
			e: Wa(e)
		};
		return {
			s: Wa(e.slice(0, r)),
			e: Wa(e.slice(r + 1))
		}
	}
	function Va(e, r) {
		if (typeof r === "undefined" || typeof r === "number") {
			return Va(e.s, e.e)
		}
		if (typeof e !== "string") e = za(e);
		if (typeof r !== "string") r = za(r);
		return e == r ? e: e + ":" + r
	}
	function $a(e) {
		var r = Ha(e);
		return "$" + La(r.s.c) + "$" + Na(r.s.r) + ":$" + La(r.e.c) + "$" + Na(r.e.r)
	}
	function Xa(e, r) {
		if (!e && !(r && r.biff <= 5 && r.biff >= 2)) throw new Error("empty sheet name");
		if (/[^\w\u4E00-\u9FFF\u3040-\u30FF]/.test(e)) return "'" + e.replace(/'/g, "''") + "'";
		return e
	}
	function Ga(e) {
		var r = {
			s: {
				c: 0,
				r: 0
			},
			e: {
				c: 0,
				r: 0
			}
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
	function ja(e, r) {
		var t = e.t == "d" && r instanceof Date;
		if (e.z != null) try {
			return e.w = ze(e.z, t ? dr(r) : r)
		} catch(a) {}
		try {
			return e.w = ze((e.XF || {}).numFmtId || (t ? 14 : 0), t ? dr(r) : r)
		} catch(a) {
			return "" + r
		}
	}
	function Ka(e, r, t) {
		if (e == null || e.t == null || e.t == "z") return "";
		if (e.w !== undefined) return e.w;
		if (e.t == "d" && !e.z && t && t.dateNF) e.z = t.dateNF;
		if (e.t == "e") return ti[e.v] || e.v;
		if (r == undefined) return ja(e, e.v);
		return ja(e, r)
	}
	function Ya(e, r) {
		var t = r && r.sheet ? r.sheet: "Sheet1";
		var a = {};
		a[t] = e;
		return {
			SheetNames: [t],
			Sheets: a
		}
	}
	function Za(e) {
		var r = {};
		var t = e || {};
		if (t.dense) r["!data"] = [];
		return r
	}
	function Ja(e, r, t) {
		var a = t || {};
		var n = e ? e["!data"] != null: a.dense;
		if (g != null && n == null) n = g;
		var i = e || {};
		if (n && !i["!data"]) i["!data"] = [];
		var s = 0,
		f = 0;
		if (i && a.origin != null) {
			if (typeof a.origin == "number") s = a.origin;
			else {
				var l = typeof a.origin == "string" ? Wa(a.origin) : a.origin;
				s = l.r;
				f = l.c
			}
			if (!i["!ref"]) i["!ref"] = "A1:A1"
		}
		var o = {
			s: {
				c: 1e7,
				r: 1e7
			},
			e: {
				c: 0,
				r: 0
			}
		};
		if (i["!ref"]) {
			var c = Ga(i["!ref"]);
			o.s.c = c.s.c;
			o.s.r = c.s.r;
			o.e.c = Math.max(o.e.c, c.e.c);
			o.e.r = Math.max(o.e.r, c.e.r);
			if (s == -1) o.e.r = s = c.e.r + 1
		}
		var u = [];
		for (var h = 0; h != r.length; ++h) {
			if (!r[h]) continue;
			if (!Array.isArray(r[h])) throw new Error("aoa_to_sheet expects an array of arrays");
			var d = s + h,
			v = "" + (d + 1);
			if (n) {
				if (!i["!data"][d]) i["!data"][d] = [];
				u = i["!data"][d]
			}
			for (var p = 0; p != r[h].length; ++p) {
				if (typeof r[h][p] === "undefined") continue;
				var m = {
					v: r[h][p]
				};
				var b = f + p;
				if (o.s.r > d) o.s.r = d;
				if (o.s.c > b) o.s.c = b;
				if (o.e.r < d) o.e.r = d;
				if (o.e.c < b) o.e.c = b;
				if (r[h][p] && typeof r[h][p] === "object" && !Array.isArray(r[h][p]) && !(r[h][p] instanceof Date)) m = r[h][p];
				else {
					if (Array.isArray(m.v)) {
						m.f = r[h][p][1];
						m.v = m.v[0]
					}
					if (m.v === null) {
						if (m.f) m.t = "n";
						else if (a.nullError) {
							m.t = "e";
							m.v = 0
						} else if (!a.sheetStubs) continue;
						else m.t = "z"
					} else if (typeof m.v === "number") m.t = "n";
					else if (typeof m.v === "boolean") m.t = "b";
					else if (m.v instanceof Date) {
						m.z = a.dateNF || q[14];
						if (!a.UTC) m.v = Dr(m.v);
						if (a.cellDates) {
							m.t = "d";
							m.w = ze(m.z, dr(m.v, a.date1904))
						} else {
							m.t = "n";
							m.v = dr(m.v, a.date1904);
							m.w = ze(m.z, m.v)
						}
					} else m.t = "s"
				}
				if (n) {
					if (u[b] && u[b].z) m.z = u[b].z;
					u[b] = m
				} else {
					var w = La(b) + v;
					if (i[w] && i[w].z) m.z = i[w].z;
					i[w] = m
				}
			}
		}
		if (o.s.c < 1e7) i["!ref"] = Va(o);
		return i
	}
	function qa(e, r) {
		return Ja(null, e, r)
	}
	function Qa(e) {
		return e._R(4, "i")
	}
	function en(e, r) {
		if (!r) r = Ea(4);
		r._W(4, e);
		return r
	}
	function rn(e) {
		var r = e._R(4);
		return r === 0 ? "": e._R(r, "dbcs")
	}
	function tn(e, r) {
		var t = false;
		if (r == null) {
			t = true;
			r = Ea(4 + 2 * e.length)
		}
		r._W(4, e.length);
		if (e.length > 0) r._W(0, e, "dbcs");
		return t ? r.slice(0, r.l) : r
	}
	function an(e) {
		return {
			ich: e._R(2),
			ifnt: e._R(2)
		}
	}
	function nn(e, r) {
		if (!r) r = Ea(4);
		r._W(2, e.ich || 0);
		r._W(2, e.ifnt || 0);
		return r
	}
	function sn(e, r) {
		var t = e.l;
		var a = e._R(1);
		var n = rn(e);
		var i = [];
		var s = {
			t: n,
			h: n
		};
		if ((a & 1) !== 0) {
			var f = e._R(4);
			for (var l = 0; l != f; ++l) i.push(an(e));
			s.r = i
		} else s.r = [{
			ich: 0,
			ifnt: 0
		}];
		e.l = t + r;
		return s
	}
	function fn(e, r) {
		var t = false;
		if (r == null) {
			t = true;
			r = Ea(15 + 4 * e.t.length)
		}
		r._W(1, 0);
		tn(e.t, r);
		return t ? r.slice(0, r.l) : r
	}
	var ln = sn;
	function on(e, r) {
		var t = false;
		if (r == null) {
			t = true;
			r = Ea(23 + 4 * e.t.length)
		}
		r._W(1, 1);
		tn(e.t, r);
		r._W(4, 1);
		nn({
			ich: 0,
			ifnt: 0
		},
		r);
		return t ? r.slice(0, r.l) : r
	}
	function cn(e) {
		var r = e._R(4);
		var t = e._R(2);
		t += e._R(1) << 16;
		e.l++;
		return {
			c: r,
			iStyleRef: t
		}
	}
	function un(e, r) {
		if (r == null) r = Ea(8);
		r._W(-4, e.c);
		r._W(3, e.iStyleRef || e.s);
		r._W(1, 0);
		return r
	}
	function hn(e) {
		var r = e._R(2);
		r += e._R(1) << 16;
		e.l++;
		return {
			c: -1,
			iStyleRef: r
		}
	}
	function dn(e, r) {
		if (r == null) r = Ea(4);
		r._W(3, e.iStyleRef || e.s);
		r._W(1, 0);
		return r
	}
	var vn = rn;
	var pn = tn;
	function mn(e) {
		var r = e._R(4);
		return r === 0 || r === 4294967295 ? "": e._R(r, "dbcs")
	}
	function bn(e, r) {
		var t = false;
		if (r == null) {
			t = true;
			r = Ea(127)
		}
		r._W(4, e.length > 0 ? e.length: 4294967295);
		if (e.length > 0) r._W(0, e, "dbcs");
		return t ? r.slice(0, r.l) : r
	}
	var gn = rn;
	var wn = mn;
	var kn = bn;
	function Tn(e) {
		var r = e.slice(e.l, e.l + 4);
		var t = r[0] & 1,
		a = r[0] & 2;
		e.l += 4;
		var n = a === 0 ? sa([0, 0, 0, 0, r[0] & 252, r[1], r[2], r[3]], 0) : da(r, 0) >> 2;
		return t ? n / 100 : n
	}
	function yn(e, r) {
		if (r == null) r = Ea(4);
		var t = 0,
		a = 0,
		n = e * 100;
		if (e == (e | 0) && e >= -(1 << 29) && e < 1 << 29) {
			a = 1
		} else if (n == (n | 0) && n >= -(1 << 29) && n < 1 << 29) {
			a = 1;
			t = 1
		}
		if (a) r._W(-4, ((t ? n: e) << 2) + (t + 2));
		else throw new Error("unsupported RkNumber " + e)
	}
	function En(e) {
		var r = {
			s: {},
			e: {}
		};
		r.s.r = e._R(4);
		r.e.r = e._R(4);
		r.s.c = e._R(4);
		r.e.c = e._R(4);
		return r
	}
	function _n(e, r) {
		if (!r) r = Ea(16);
		r._W(4, e.s.r);
		r._W(4, e.e.r);
		r._W(4, e.s.c);
		r._W(4, e.e.c);
		return r
	}
	var Sn = En;
	var xn = _n;
	function An(e) {
		if (e.length - e.l < 8) throw "XLS Xnum Buffer underflow";
		return e._R(8, "f")
	}
	function Cn(e, r) {
		return (r || Ea(8))._W(8, e, "f")
	}
	function Rn(e) {
		var r = {};
		var t = e._R(1);
		var a = t >>> 1;
		var n = e._R(1);
		var i = e._R(2, "i");
		var s = e._R(1);
		var f = e._R(1);
		var l = e._R(1);
		e.l++;
		switch (a) {
		case 0:
			r.auto = 1;
			break;
		case 1:
			r.index = n;
			var o = ri[n];
			if (o) r.rgb = Xo(o);
			break;
		case 2:
			r.rgb = Xo([s, f, l]);
			break;
		case 3:
			r.theme = n;
			break;
		}
		if (i != 0) r.tint = i > 0 ? i / 32767 : i / 32768;
		return r
	}
	function On(e, r) {
		if (!r) r = Ea(8);
		if (!e || e.auto) {
			r._W(4, 0);
			r._W(4, 0);
			return r
		}
		if (e.index != null) {
			r._W(1, 2);
			r._W(1, e.index)
		} else if (e.theme != null) {
			r._W(1, 6);
			r._W(1, e.theme)
		} else {
			r._W(1, 5);
			r._W(1, 0)
		}
		var t = e.tint || 0;
		if (t > 0) t *= 32767;
		else if (t < 0) t *= 32768;
		r._W(2, t);
		if (!e.rgb || e.theme != null) {
			r._W(2, 0);
			r._W(1, 0);
			r._W(1, 0)
		} else {
			var a = e.rgb || "FFFFFF";
			if (typeof a == "number") a = ("000000" + a.toString(16)).slice(-6);
			r._W(1, parseInt(a.slice(0, 2), 16));
			r._W(1, parseInt(a.slice(2, 4), 16));
			r._W(1, parseInt(a.slice(4, 6), 16));
			r._W(1, 255)
		}
		return r
	}
	function In(e) {
		var r = e._R(1);
		e.l++;
		var t = {
			fBold: r & 1,
			fItalic: r & 2,
			fUnderline: r & 4,
			fStrikeout: r & 8,
			fOutline: r & 16,
			fShadow: r & 32,
			fCondense: r & 64,
			fExtend: r & 128
		};
		return t
	}
	function Nn(e, r) {
		if (!r) r = Ea(2);
		var t = (e.italic ? 2 : 0) | (e.strike ? 8 : 0) | (e.outline ? 16 : 0) | (e.shadow ? 32 : 0) | (e.condense ? 64 : 0) | (e.extend ? 128 : 0);
		r._W(1, t);
		r._W(1, 0);
		return r
	}
	function Fn(e, r) {
		var t = {
			2 : "BITMAP",
			3 : "METAFILEPICT",
			8 : "DIB",
			14 : "ENHMETAFILE"
		};
		var a = e._R(4);
		switch (a) {
		case 0:
			return "";
		case 4294967295:
			;
		case 4294967294:
			return t[e._R(4)] || "";
		}
		if (a > 400) throw new Error("Unsupported Clipboard: " + a.toString(16));
		e.l -= 4;
		return e._R(0, r == 1 ? "lpstr": "lpwstr")
	}
	function Dn(e) {
		return Fn(e, 1)
	}
	function Pn(e) {
		return Fn(e, 2)
	}
	var Ln = 2;
	var Mn = 3;
	var Un = 11;
	var Bn = 12;
	var Wn = 19;
	var zn = 64;
	var Hn = 65;
	var Vn = 71;
	var $n = 4108;
	var Xn = 4126;
	var Gn = 80;
	var jn = 81;
	var Kn = [Gn, jn];
	var Yn = {
		1 : {
			n: "CodePage",
			t: Ln
		},
		2 : {
			n: "Category",
			t: Gn
		},
		3 : {
			n: "PresentationFormat",
			t: Gn
		},
		4 : {
			n: "ByteCount",
			t: Mn
		},
		5 : {
			n: "LineCount",
			t: Mn
		},
		6 : {
			n: "ParagraphCount",
			t: Mn
		},
		7 : {
			n: "SlideCount",
			t: Mn
		},
		8 : {
			n: "NoteCount",
			t: Mn
		},
		9 : {
			n: "HiddenCount",
			t: Mn
		},
		10 : {
			n: "MultimediaClipCount",
			t: Mn
		},
		11 : {
			n: "ScaleCrop",
			t: Un
		},
		12 : {
			n: "HeadingPairs",
			t: $n
		},
		13 : {
			n: "TitlesOfParts",
			t: Xn
		},
		14 : {
			n: "Manager",
			t: Gn
		},
		15 : {
			n: "Company",
			t: Gn
		},
		16 : {
			n: "LinksUpToDate",
			t: Un
		},
		17 : {
			n: "CharacterCount",
			t: Mn
		},
		19 : {
			n: "SharedDoc",
			t: Un
		},
		22 : {
			n: "HyperlinksChanged",
			t: Un
		},
		23 : {
			n: "AppVersion",
			t: Mn,
			p: "version"
		},
		24 : {
			n: "DigSig",
			t: Hn
		},
		26 : {
			n: "ContentType",
			t: Gn
		},
		27 : {
			n: "ContentStatus",
			t: Gn
		},
		28 : {
			n: "Language",
			t: Gn
		},
		29 : {
			n: "Version",
			t: Gn
		},
		255 : {},
		2147483648 : {
			n: "Locale",
			t: Wn
		},
		2147483651 : {
			n: "Behavior",
			t: Wn
		},
		1919054434 : {}
	};
	var Zn = {
		1 : {
			n: "CodePage",
			t: Ln
		},
		2 : {
			n: "Title",
			t: Gn
		},
		3 : {
			n: "Subject",
			t: Gn
		},
		4 : {
			n: "Author",
			t: Gn
		},
		5 : {
			n: "Keywords",
			t: Gn
		},
		6 : {
			n: "Comments",
			t: Gn
		},
		7 : {
			n: "Template",
			t: Gn
		},
		8 : {
			n: "LastAuthor",
			t: Gn
		},
		9 : {
			n: "RevNumber",
			t: Gn
		},
		10 : {
			n: "EditTime",
			t: zn
		},
		11 : {
			n: "LastPrinted",
			t: zn
		},
		12 : {
			n: "CreatedDate",
			t: zn
		},
		13 : {
			n: "ModifiedDate",
			t: zn
		},
		14 : {
			n: "PageCount",
			t: Mn
		},
		15 : {
			n: "WordCount",
			t: Mn
		},
		16 : {
			n: "CharCount",
			t: Mn
		},
		17 : {
			n: "Thumbnail",
			t: Vn
		},
		18 : {
			n: "Application",
			t: Gn
		},
		19 : {
			n: "DocSecurity",
			t: Mn
		},
		255 : {},
		2147483648 : {
			n: "Locale",
			t: Wn
		},
		2147483651 : {
			n: "Behavior",
			t: Wn
		},
		1919054434 : {}
	};
	var Jn = {
		1 : "US",
		2 : "CA",
		3 : "",
		7 : "RU",
		20 : "EG",
		30 : "GR",
		31 : "NL",
		32 : "BE",
		33 : "FR",
		34 : "ES",
		36 : "HU",
		39 : "IT",
		41 : "CH",
		43 : "AT",
		44 : "GB",
		45 : "DK",
		46 : "SE",
		47 : "NO",
		48 : "PL",
		49 : "DE",
		52 : "MX",
		55 : "BR",
		61 : "AU",
		64 : "NZ",
		66 : "TH",
		81 : "JP",
		82 : "KR",
		84 : "VN",
		86 : "CN",
		90 : "TR",
		105 : "JS",
		213 : "DZ",
		216 : "MA",
		218 : "LY",
		351 : "PT",
		354 : "IS",
		358 : "FI",
		420 : "CZ",
		886 : "TW",
		961 : "LB",
		962 : "JO",
		963 : "SY",
		964 : "IQ",
		965 : "KW",
		966 : "SA",
		971 : "AE",
		972 : "IL",
		974 : "QA",
		981 : "IR",
		65535 : "US"
	};
	var qn = [null, "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"];
	function Qn(e) {
		return e.map(function(e) {
			return [e >> 16 & 255, e >> 8 & 255, e & 255]
		})
	}
	var ei = Qn([0, 16777215, 16711680, 65280, 255, 16776960, 16711935, 65535, 0, 16777215, 16711680, 65280, 255, 16776960, 16711935, 65535, 8388608, 32768, 128, 8421376, 8388736, 32896, 12632256, 8421504, 10066431, 10040166, 16777164, 13434879, 6684774, 16744576, 26316, 13421823, 128, 16711935, 16776960, 65535, 8388736, 8388608, 32896, 255, 52479, 13434879, 13434828, 16777113, 10079487, 16751052, 13408767, 16764057, 3368703, 3394764, 10079232, 16763904, 16750848, 16737792, 6710937, 9868950, 13158, 3381606, 13056, 3355392, 10040064, 10040166, 3355545, 3355443, 0, 16777215, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
	var ri = Tr(ei);
	var ti = {
		0 : "#NULL!",
		7 : "#DIV/0!",
		15 : "#VALUE!",
		23 : "#REF!",
		29 : "#NAME?",
		36 : "#NUM!",
		42 : "#N/A",
		43 : "#GETTING_DATA",
		255 : "#WTF?"
	};
	var ai = {
		"#NULL!": 0,
		"#DIV/0!": 7,
		"#VALUE!": 15,
		"#REF!": 23,
		"#NAME?": 29,
		"#NUM!": 36,
		"#N/A": 42,
		"#GETTING_DATA": 43,
		"#WTF?": 255
	};
	var ni = ["_xlnm.Consolidate_Area", "_xlnm.Auto_Open", "_xlnm.Auto_Close", "_xlnm.Extract", "_xlnm.Database", "_xlnm.Criteria", "_xlnm.Print_Area", "_xlnm.Print_Titles", "_xlnm.Recorder", "_xlnm.Data_Form", "_xlnm.Auto_Activate", "_xlnm.Auto_Deactivate", "_xlnm.Sheet_Title", "_xlnm._FilterDatabase"];
	var ii = {
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
		sheet: "js"
	};
	var si = {
		workbooks: {
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
			xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
			xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
			xlam: "application/vnd.ms-excel.addin.macroEnabled.main+xml",
			xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
		},
		strs: {
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
			xlsb: "application/vnd.ms-excel.sharedStrings"
		},
		comments: {
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
			xlsb: "application/vnd.ms-excel.comments"
		},
		sheets: {
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
			xlsb: "application/vnd.ms-excel.worksheet"
		},
		charts: {
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
			xlsb: "application/vnd.ms-excel.chartsheet"
		},
		dialogs: {
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
			xlsb: "application/vnd.ms-excel.dialogsheet"
		},
		macros: {
			xlsx: "application/vnd.ms-excel.macrosheet+xml",
			xlsb: "application/vnd.ms-excel.macrosheet"
		},
		metadata: {
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
			xlsb: "application/vnd.ms-excel.sheetMetadata"
		},
		styles: {
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
			xlsb: "application/vnd.ms-excel.styles"
		}
	};
	function fi() {
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
			xmlns: ""
		}
	}
	function li(e) {
		var r = fi();
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
				if (r[ii[a.ContentType]] !== undefined) r[ii[a.ContentType]].push(a.PartName);
				break;
			}
		});
		if (r.xmlns !== Lt.CT) throw new Error("Unknown Namespace: " + r.xmlns);
		r.calcchain = r.calcchains.length > 0 ? r.calcchains[0] : "";
		r.sst = r.strs.length > 0 ? r.strs[0] : "";
		r.style = r.styles.length > 0 ? r.styles[0] : "";
		r.defaults = t;
		delete r.calcchains;
		return r
	}
	function oi(e, r, t) {
		var a = or(ii);
		var n = [],
		i;
		if (!t) {
			n[n.length] = Kr;
			n[n.length] = It("Types", null, {
				xmlns: Lt.CT,
				"xmlns:xsd": Lt.xsd,
				"xmlns:xsi": Lt.xsi
			});
			n = n.concat([["xml", "application/xml"], ["bin", "application/vnd.ms-excel.sheet.binary.macroEnabled.main"], ["vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"], ["data", "application/vnd.openxmlformats-officedocument.model+data"], ["bmp", "image/bmp"], ["png", "image/png"], ["gif", "image/gif"], ["emf", "image/x-emf"], ["wmf", "image/x-wmf"], ["jpg", "image/jpeg"], ["jpeg", "image/jpeg"], ["tif", "image/tiff"], ["tiff", "image/tiff"], ["pdf", "application/pdf"], ["rels", "application/vnd.openxmlformats-package.relationships+xml"]].map(function(e) {
				return It("Default", null, {
					Extension: e[0],
					ContentType: e[1]
				})
			}))
		}
		var s = function(t) {
			if (e[t] && e[t].length > 0) {
				i = e[t][0];
				n[n.length] = It("Override", null, {
					PartName: (i[0] == "/" ? "": "/") + i,
					ContentType: si[t][r.bookType] || si[t]["xlsx"]
				})
			}
		};
		var f = function(t) { (e[t] || []).forEach(function(e) {
				n[n.length] = It("Override", null, {
					PartName: (e[0] == "/" ? "": "/") + e,
					ContentType: si[t][r.bookType] || si[t]["xlsx"]
				})
			})
		};
		var l = function(r) { (e[r] || []).forEach(function(e) {
				n[n.length] = It("Override", null, {
					PartName: (e[0] == "/" ? "": "/") + e,
					ContentType: a[r][0]
				})
			})
		};
		s("workbooks");
		f("sheets");
		f("charts");
		l("themes"); ["strs", "styles"].forEach(s); ["coreprops", "extprops", "custprops"].forEach(l);
		l("vba");
		l("comments");
		l("threadedcomments");
		l("drawings");
		f("metadata");
		l("people");
		if (!t && n.length > 2) {
			n[n.length] = "</Types>";
			n[1] = n[1].replace("/>", ">")
		}
		return n.join("")
	}
	var ci = {
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
		WS: ["http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"],
		DS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
		MS: "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
		IMG: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
		DRAW: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
		XLMETA: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
		TCMNT: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
		PEOPLE: "http://schemas.microsoft.com/office/2017/10/relationships/person",
		CONN: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
		VBA: "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
	};
	function ui(e) {
		var r = e.lastIndexOf("/");
		return e.slice(0, r + 1) + "_rels/" + e.slice(r + 1) + ".rels"
	}
	function hi(e, r) {
		var t = {
			"!id": {}
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
				var s = n.TargetMode === "External" ? n.Target: jr(n.Target, r);
				t[s] = i;
				a[n.Id] = i
			}
		});
		t["!id"] = a;
		return t
	}
	function di(e) {
		var r = [Kr, It("Relationships", null, {
			xmlns: Lt.RELS
		})];
		ir(e["!id"]).forEach(function(t) {
			r[r.length] = It("Relationship", null, e["!id"][t])
		});
		if (r.length > 2) {
			r[r.length] = "</Relationships>";
			r[1] = r[1].replace("/>", ">")
		}
		return r.join("")
	}
	function vi(e, r, t, a, n, i) {
		if (!n) n = {};
		if (!e["!id"]) e["!id"] = {};
		if (!e["!idx"]) e["!idx"] = 1;
		if (r < 0) for (r = e["!idx"]; e["!id"]["rId" + r]; ++r) {}
		e["!idx"] = r + 1;
		n.Id = "rId" + r;
		n.Type = a;
		n.Target = t;
		if (i) n.TargetMode = i;
		else if ([ci.HLINK, ci.XPATH, ci.XMISS].indexOf(n.Type) > -1) n.TargetMode = "External";
		if (e["!id"][n.Id]) throw new Error("Cannot rewrite rId " + r);
		e["!id"][n.Id] = n;
		e[("/" + n.Target).replace("//", "/")] = n;
		return r
	}
	var pi = "application/vnd.oasis.opendocument.spreadsheet";
	function mi(e, r) {
		var t = Dt(e);
		var a;
		var n;
		while (a = Pt.exec(t)) switch (a[3]) {
		case "manifest":
			break;
		case "file-entry":
			n = rt(a[0], false);
			if (n.path == "/" && n.type !== pi) throw new Error("This OpenDocument is not a spreadsheet");
			break;
		case "encryption-data":
			;
		case "algorithm":
			;
		case "start-key-generation":
			;
		case "key-derivation":
			throw new Error("Unsupported ODS Encryption");
		default:
			if (r && r.WTF) throw a;
		}
	}
	function bi(e) {
		var r = [Kr];
		r.push('<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">\n');
		r.push('  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>\n');
		for (var t = 0; t < e.length; ++t) r.push('  <manifest:file-entry manifest:full-path="' + e[t][0] + '" manifest:media-type="' + e[t][1] + '"/>\n');
		r.push("</manifest:manifest>");
		return r.join("")
	}
	function gi(e, r, t) {
		return ['  <rdf:Description rdf:about="' + e + '">\n', '    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/' + (t || "odf") + "#" + r + '"/>\n', "  </rdf:Description>\n"].join("")
	}
	function wi(e, r) {
		return ['  <rdf:Description rdf:about="' + e + '">\n', '    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="' + r + '"/>\n', "  </rdf:Description>\n"].join("")
	}
	function ki(e) {
		var r = [Kr];
		r.push('<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">\n');
		for (var t = 0; t != e.length; ++t) {
			r.push(gi(e[t][0], e[t][1]));
			r.push(wi("", e[t][0]))
		}
		r.push(gi("", "Document", "pkg"));
		r.push("</rdf:RDF>");
		return r.join("")
	}
	function Ti(r, t) {
		return '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.2"><office:meta><meta:generator>SheetJS ' + e.version + "</meta:generator></office:meta></office:document-meta>"
	}
	var yi = [["cp:category", "Category"], ["cp:contentStatus", "ContentStatus"], ["cp:keywords", "Keywords"], ["cp:lastModifiedBy", "LastAuthor"], ["cp:lastPrinted", "LastPrinted"], ["cp:revision", "RevNumber"], ["cp:version", "Version"], ["dc:creator", "Author"], ["dc:description", "Comments"], ["dc:identifier", "Identifier"], ["dc:language", "Language"], ["dc:subject", "Subject"], ["dc:title", "Title"], ["dcterms:created", "CreatedDate", "date"], ["dcterms:modified", "ModifiedDate", "date"]];
	var Ei = function() {
		var e = new Array(yi.length);
		for (var r = 0; r < yi.length; ++r) {
			var t = yi[r];
			var a = "(?:" + t[0].slice(0, t[0].indexOf(":")) + ":)" + t[0].slice(t[0].indexOf(":") + 1);
			e[r] = new RegExp("<" + a + "[^>]*>([\\s\\S]*?)</" + a + ">")
		}
		return e
	} ();
	function _i(e) {
		var r = {};
		e = kt(e);
		for (var t = 0; t < yi.length; ++t) {
			var a = yi[t],
			n = e.match(Ei[t]);
			if (n != null && n.length > 0) r[a[1]] = it(n[1]);
			if (a[2] === "date" && r[a[1]]) r[a[1]] = wr(r[a[1]])
		}
		return r
	}
	function Si(e, r, t, a, n) {
		if (n[e] != null || r == null || r === "") return;
		n[e] = r;
		r = lt(r);
		a[a.length] = t ? It(e, r, t) : Rt(e, r)
	}
	function xi(e, r) {
		var t = r || {};
		var a = [Kr, It("cp:coreProperties", null, {
			"xmlns:cp": Lt.CORE_PROPS,
			"xmlns:dc": Lt.dc,
			"xmlns:dcterms": Lt.dcterms,
			"xmlns:dcmitype": Lt.dcmitype,
			"xmlns:xsi": Lt.xsi
		})],
		n = {};
		if (!e && !t.Props) return a.join("");
		if (e) {
			if (e.CreatedDate != null) Si("dcterms:created", typeof e.CreatedDate === "string" ? e.CreatedDate: Nt(e.CreatedDate, t.WTF), {
				"xsi:type": "dcterms:W3CDTF"
			},
			a, n);
			if (e.ModifiedDate != null) Si("dcterms:modified", typeof e.ModifiedDate === "string" ? e.ModifiedDate: Nt(e.ModifiedDate, t.WTF), {
				"xsi:type": "dcterms:W3CDTF"
			},
			a, n)
		}
		for (var i = 0; i != yi.length; ++i) {
			var s = yi[i];
			var f = t.Props && t.Props[s[1]] != null ? t.Props[s[1]] : e ? e[s[1]] : null;
			if (f === true) f = "1";
			else if (f === false) f = "0";
			else if (typeof f == "number") f = String(f);
			if (f != null) Si(s[0], f, null, a, n)
		}
		if (a.length > 2) {
			a[a.length] = "</cp:coreProperties>";
			a[1] = a[1].replace("/>", ">")
		}
		return a.join("")
	}
	var Ai = [["Application", "Application", "string"], ["AppVersion", "AppVersion", "string"], ["Company", "Company", "string"], ["DocSecurity", "DocSecurity", "string"], ["Manager", "Manager", "string"], ["HyperlinksChanged", "HyperlinksChanged", "bool"], ["SharedDoc", "SharedDoc", "bool"], ["LinksUpToDate", "LinksUpToDate", "bool"], ["ScaleCrop", "ScaleCrop", "bool"], ["HeadingPairs", "HeadingPairs", "raw"], ["TitlesOfParts", "TitlesOfParts", "raw"]];
	var Ci = ["Worksheets", "SheetNames", "NamedRanges", "DefinedNames", "Chartsheets", "ChartNames"];
	function Ri(e, r, t, a) {
		var n = [];
		if (typeof e == "string") n = At(e, a);
		else for (var i = 0; i < e.length; ++i) n = n.concat(e[i].map(function(e) {
			return {
				v: e
			}
		}));
		var s = typeof r == "string" ? At(r, a).map(function(e) {
			return e.v
		}) : r;
		var f = 0,
		l = 0;
		if (s.length > 0) for (var o = 0; o !== n.length; o += 2) {
			l = +n[o + 1].v;
			switch (n[o].v) {
			case "Worksheets":
				;
			case "工作表":
				;
			case "Листы":
				;
			case "أوراق العمل":
				;
			case "ワークシート":
				;
			case "גליונות עבודה":
				;
			case "Arbeitsblätter":
				;
			case "Çalışma Sayfaları":
				;
			case "Feuilles de calcul":
				;
			case "Fogli di lavoro":
				;
			case "Folhas de cálculo":
				;
			case "Planilhas":
				;
			case "Regneark":
				;
			case "Hojas de cálculo":
				;
			case "Werkbladen":
				t.Worksheets = l;
				t.SheetNames = s.slice(f, f + l);
				break;
			case "Named Ranges":
				;
			case "Rangos con nombre":
				;
			case "名前付き一覧":
				;
			case "Benannte Bereiche":
				;
			case "Navngivne områder":
				t.NamedRanges = l;
				t.DefinedNames = s.slice(f, f + l);
				break;
			case "Charts":
				;
			case "Diagramme":
				t.Chartsheets = l;
				t.ChartNames = s.slice(f, f + l);
				break;
			}
			f += l
		}
	}
	function Oi(e, r, t) {
		var a = {};
		if (!r) r = {};
		e = kt(e);
		Ai.forEach(function(t) {
			var n = (e.match(yt(t[0])) || [])[1];
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
				break;
			}
		});
		if (a.HeadingPairs && a.TitlesOfParts) Ri(a.HeadingPairs, a.TitlesOfParts, r, t);
		return r
	}
	function Ii(e) {
		var r = [],
		t = It;
		if (!e) e = {};
		e.Application = "SheetJS";
		r[r.length] = Kr;
		r[r.length] = It("Properties", null, {
			xmlns: Lt.EXT_PROPS,
			"xmlns:vt": Lt.vt
		});
		Ai.forEach(function(a) {
			if (e[a[1]] === undefined) return;
			var n;
			switch (a[2]) {
			case "string":
				n = lt(String(e[a[1]]));
				break;
			case "bool":
				n = e[a[1]] ? "true": "false";
				break;
			}
			if (n !== undefined) r[r.length] = t(a[0], n)
		});
		r[r.length] = t("HeadingPairs", t("vt:vector", t("vt:variant", "<vt:lpstr>Worksheets</vt:lpstr>") + t("vt:variant", t("vt:i4", String(e.Worksheets))), {
			size: 2,
			baseType: "variant"
		}));
		r[r.length] = t("TitlesOfParts", t("vt:vector", e.SheetNames.map(function(e) {
			return "<vt:lpstr>" + lt(e) + "</vt:lpstr>"
		}).join(""), {
			size: e.Worksheets,
			baseType: "lpstr"
		}));
		if (r.length > 2) {
			r[r.length] = "</Properties>";
			r[1] = r[1].replace("/>", ">")
		}
		return r.join("")
	}
	var Ni = /<[^>]+>[^<]*/g;
	function Fi(e, r) {
		var t = {},
		a = "";
		var n = e.match(Ni);
		if (n) for (var i = 0; i != n.length; ++i) {
			var s = n[i],
			f = rt(s);
			switch (tt(f[0])) {
			case "<?xml":
				break;
			case "<Properties":
				break;
			case "<property":
				a = it(f.name);
				break;
			case "</property>":
				a = null;
				break;
			default:
				if (s.indexOf("<vt:") === 0) {
					var l = s.split(">");
					var o = l[0].slice(4),
					c = l[1];
					switch (o) {
					case "lpstr":
						;
					case "bstr":
						;
					case "lpwstr":
						t[a] = it(c);
						break;
					case "bool":
						t[a] = pt(c);
						break;
					case "i1":
						;
					case "i2":
						;
					case "i4":
						;
					case "i8":
						;
					case "int":
						;
					case "uint":
						t[a] = parseInt(c, 10);
						break;
					case "r4":
						;
					case "r8":
						;
					case "decimal":
						t[a] = parseFloat(c);
						break;
					case "filetime":
						;
					case "date":
						t[a] = wr(c);
						break;
					case "cy":
						;
					case "error":
						t[a] = it(c);
						break;
					default:
						if (o.slice(-1) == "/") break;
						if (r.WTF && typeof console !== "undefined") console.warn("Unexpected", s, o, l);
					}
				} else if (s.slice(0, 2) === "</") {} else if (r.WTF) throw new Error(s);
			}
		}
		return t
	}
	function Di(e) {
		var r = [Kr, It("Properties", null, {
			xmlns: Lt.CUST_PROPS,
			"xmlns:vt": Lt.vt
		})];
		if (!e) return r.join("");
		var t = 1;
		ir(e).forEach(function a(n) {++t;
			r[r.length] = It("property", Ft(e[n], true), {
				fmtid: "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
				pid: t,
				name: lt(n)
			})
		});
		if (r.length > 2) {
			r[r.length] = "</Properties>";
			r[1] = r[1].replace("/>", ">")
		}
		return r.join("")
	}
	var Pi = {
		Title: "Title",
		Subject: "Subject",
		Author: "Author",
		Keywords: "Keywords",
		Comments: "Description",
		LastAuthor: "LastAuthor",
		RevNumber: "Revision",
		Application: "AppName",
		LastPrinted: "LastPrinted",
		CreatedDate: "Created",
		ModifiedDate: "LastSaved",
		Category: "Category",
		Manager: "Manager",
		Company: "Company",
		AppVersion: "Version",
		ContentStatus: "ContentStatus",
		Identifier: "Identifier",
		Language: "Language"
	};
	var Li;
	function Mi(e, r, t) {
		if (!Li) Li = fr(Pi);
		r = Li[r] || r;
		e[r] = t
	}
	function Ui(e, r) {
		var t = [];
		ir(Pi).map(function(e) {
			for (var r = 0; r < yi.length; ++r) if (yi[r][1] == e) return yi[r];
			for (r = 0; r < Ai.length; ++r) if (Ai[r][1] == e) return Ai[r];
			throw e
		}).forEach(function(a) {
			if (e[a[1]] == null) return;
			var n = r && r.Props && r.Props[a[1]] != null ? r.Props[a[1]] : e[a[1]];
			switch (a[2]) {
			case "date":
				n = new Date(n).toISOString().replace(/\.\d*Z/, "Z");
				break;
			}
			if (typeof n == "number") n = String(n);
			else if (n === true || n === false) {
				n = n ? "1": "0"
			} else if (n instanceof Date) n = new Date(n).toISOString().replace(/\.\d*Z/, "");
			t.push(Rt(Pi[a[1]] || a[1], n))
		});
		return It("DocumentProperties", t.join(""), {
			xmlns: Ut.o
		})
	}
	function Bi(e, r) {
		var t = ["Worksheets", "SheetNames"];
		var a = "CustomDocumentProperties";
		var n = [];
		if (e) ir(e).forEach(function(r) {
			if (!Object.prototype.hasOwnProperty.call(e, r)) return;
			for (var a = 0; a < yi.length; ++a) if (r == yi[a][1]) return;
			for (a = 0; a < Ai.length; ++a) if (r == Ai[a][1]) return;
			for (a = 0; a < t.length; ++a) if (r == t[a]) return;
			var i = e[r];
			var s = "string";
			if (typeof i == "number") {
				s = "float";
				i = String(i)
			} else if (i === true || i === false) {
				s = "boolean";
				i = i ? "1": "0"
			} else i = String(i);
			n.push(It(ot(r), i, {
				"dt:dt": s
			}))
		});
		if (r) ir(r).forEach(function(t) {
			if (!Object.prototype.hasOwnProperty.call(r, t)) return;
			if (e && Object.prototype.hasOwnProperty.call(e, t)) return;
			var a = r[t];
			var i = "string";
			if (typeof a == "number") {
				i = "float";
				a = String(a)
			} else if (a === true || a === false) {
				i = "boolean";
				a = a ? "1": "0"
			} else if (a instanceof Date) {
				i = "dateTime.tz";
				a = a.toISOString()
			} else a = String(a);
			n.push(It(ot(t), a, {
				"dt:dt": i
			}))
		});
		return "<" + a + ' xmlns="' + Ut.o + '">' + n.join("") + "</" + a + ">"
	}
	function Wi(e) {
		var r = e._R(4),
		t = e._R(4);
		return new Date((t / 1e7 * Math.pow(2, 32) + r / 1e7 - 11644473600) * 1e3).toISOString().replace(/\.000/, "")
	}
	function zi(e) {
		var r = typeof e == "string" ? new Date(Date.parse(e)) : e;
		var t = r.getTime() / 1e3 + 11644473600;
		var a = t % Math.pow(2, 32),
		n = (t - a) / Math.pow(2, 32);
		a *= 1e7;
		n *= 1e7;
		var i = a / Math.pow(2, 32) | 0;
		if (i > 0) {
			a = a % Math.pow(2, 32);
			n += i
		}
		var s = Ea(8);
		s._W(4, a);
		s._W(4, n);
		return s
	}
	function Hi(e, r, t) {
		var a = e.l;
		var n = e._R(0, "lpstr-cp");
		if (t) while (e.l - a & 3)++e.l;
		return n
	}
	function Vi(e, r, t) {
		var a = e._R(0, "lpwstr");
		if (t) e.l += 4 - (a.length + 1 & 3) & 3;
		return a
	}
	function $i(e, r, t) {
		if (r === 31) return Vi(e);
		return Hi(e, r, t)
	}
	function Xi(e, r, t) {
		return $i(e, r, t === false ? 0 : 4)
	}
	function Gi(e, r) {
		if (!r) throw new Error("VtUnalignedString must have positive length");
		return $i(e, r, 0)
	}
	function ji(e) {
		var r = e._R(4);
		var t = [];
		for (var a = 0; a != r; ++a) {
			var n = e.l;
			t[a] = e._R(0, "lpwstr").replace(M, "");
			if (e.l - n & 2) e.l += 2
		}
		return t
	}
	function Ki(e) {
		var r = e._R(4);
		var t = [];
		for (var a = 0; a != r; ++a) t[a] = e._R(0, "lpstr-cp").replace(M, "");
		return t
	}
	function Yi(e) {
		var r = e.l;
		var t = es(e, jn);
		if (e[e.l] == 0 && e[e.l + 1] == 0 && e.l - r & 2) e.l += 2;
		var a = es(e, Mn);
		return [t, a]
	}
	function Zi(e) {
		var r = e._R(4);
		var t = [];
		for (var a = 0; a < r / 2; ++a) t.push(Yi(e));
		return t
	}
	function Ji(e, r) {
		var t = e._R(4);
		var a = {};
		for (var n = 0; n != t; ++n) {
			var i = e._R(4);
			var s = e._R(4);
			a[i] = e._R(s, r === 1200 ? "utf16le": "utf8").replace(M, "").replace(U, "!");
			if (r === 1200 && s % 2) e.l += 2
		}
		if (e.l & 3) e.l = e.l >> 2 + 1 << 2;
		return a
	}
	function qi(e) {
		var r = e._R(4);
		var t = e.slice(e.l, e.l + r);
		e.l += r;
		if ((r & 3) > 0) e.l += 4 - (r & 3) & 3;
		return t
	}
	function Qi(e) {
		var r = {};
		r.Size = e._R(4);
		e.l += r.Size + 3 - (r.Size - 1) % 4;
		return r
	}
	function es(e, r, t) {
		var a = e._R(2),
		n,
		i = t || {};
		e.l += 2;
		if (r !== Bn) if (a !== r && Kn.indexOf(r) === -1 && !((r & 65534) == 4126 && (a & 65534) == 4126)) throw new Error("Expected type " + r + " saw " + a);
		switch (r === Bn ? a: r) {
		case 2:
			n = e._R(2, "i");
			if (!i.raw) e.l += 2;
			return n;
		case 3:
			n = e._R(4, "i");
			return n;
		case 11:
			return e._R(4) !== 0;
		case 19:
			n = e._R(4);
			return n;
		case 30:
			return Hi(e, a, 4).replace(M, "");
		case 31:
			return Vi(e);
		case 64:
			return Wi(e);
		case 65:
			return qi(e);
		case 71:
			return Qi(e);
		case 80:
			return Xi(e, a, !i.raw).replace(M, "");
		case 81:
			return Gi(e, a).replace(M, "");
		case 4108:
			return Zi(e);
		case 4126:
			;
		case 4127:
			return a == 4127 ? ji(e) : Ki(e);
		default:
			throw new Error("TypedPropertyValue unrecognized type " + r + " " + a);
		}
	}
	function rs(e, r) {
		var t = Ea(4),
		a = Ea(4);
		t._W(4, e == 80 ? 31 : e);
		switch (e) {
		case 3:
			a._W(-4, r);
			break;
		case 5:
			a = Ea(8);
			a._W(8, r, "f");
			break;
		case 11:
			a._W(4, r ? 1 : 0);
			break;
		case 64:
			a = zi(r);
			break;
		case 31:
			;
		case 80:
			a = Ea(4 + 2 * (r.length + 1) + (r.length % 2 ? 0 : 2));
			a._W(4, r.length + 1);
			a._W(0, r, "dbcs");
			while (a.l != a.length) a._W(1, 0);
			break;
		default:
			throw new Error("TypedPropertyValue unrecognized type " + e + " " + r);;
		}
		return P([t, a])
	}
	function ts(e, r) {
		var t = e.l;
		var a = e._R(4);
		var n = e._R(4);
		var i = [],
		s = 0;
		var f = 0;
		var o = -1,
		c = {};
		for (s = 0; s != n; ++s) {
			var u = e._R(4);
			var h = e._R(4);
			i[s] = [u, h + t]
		}
		i.sort(function(e, r) {
			return e[1] - r[1]
		});
		var d = {};
		for (s = 0; s != n; ++s) {
			if (e.l !== i[s][1]) {
				var v = true;
				if (s > 0 && r) switch (r[i[s - 1][0]].t) {
				case 2:
					if (e.l + 2 === i[s][1]) {
						e.l += 2;
						v = false
					}
					break;
				case 80:
					if (e.l <= i[s][1]) {
						e.l = i[s][1];
						v = false
					}
					break;
				case 4108:
					if (e.l <= i[s][1]) {
						e.l = i[s][1];
						v = false
					}
					break;
				}
				if ((!r || s == 0) && e.l <= i[s][1]) {
					v = false;
					e.l = i[s][1]
				}
				if (v) throw new Error("Read Error: Expected address " + i[s][1] + " at " + e.l + " :" + s)
			}
			if (r) {
				if (i[s][0] == 0 && i.length > s + 1 && i[s][1] == i[s + 1][1]) continue;
				var p = r[i[s][0]];
				d[p.n] = es(e, p.t, {
					raw: true
				});
				if (p.p === "version") d[p.n] = String(d[p.n] >> 16) + "." + ("0000" + String(d[p.n] & 65535)).slice(-4);
				if (p.n == "CodePage") switch (d[p.n]) {
				case 0:
					d[p.n] = 1252;
				case 874:
					;
				case 932:
					;
				case 936:
					;
				case 949:
					;
				case 950:
					;
				case 1250:
					;
				case 1251:
					;
				case 1253:
					;
				case 1254:
					;
				case 1255:
					;
				case 1256:
					;
				case 1257:
					;
				case 1258:
					;
				case 1e4:
					;
				case 1200:
					;
				case 1201:
					;
				case 1252:
					;
				case 65e3:
					;
				case - 536 : ;
				case 65001:
					;
				case - 535 : l(f = d[p.n] >>> 0 & 65535);
					break;
				default:
					throw new Error("Unsupported CodePage: " + d[p.n]);
				}
			} else {
				if (i[s][0] === 1) {
					f = d.CodePage = es(e, Ln);
					l(f);
					if (o !== -1) {
						var m = e.l;
						e.l = i[o][1];
						c = Ji(e, f);
						e.l = m
					}
				} else if (i[s][0] === 0) {
					if (f === 0) {
						o = s;
						e.l = i[s + 1][1];
						continue
					}
					c = Ji(e, f)
				} else {
					var b = c[i[s][0]];
					var g;
					switch (e[e.l]) {
					case 65:
						e.l += 4;
						g = qi(e);
						break;
					case 30:
						e.l += 4;
						g = Xi(e, e[e.l - 4]).replace(/\u0000+$/, "");
						break;
					case 31:
						e.l += 4;
						g = Xi(e, e[e.l - 4]).replace(/\u0000+$/, "");
						break;
					case 3:
						e.l += 4;
						g = e._R(4, "i");
						break;
					case 19:
						e.l += 4;
						g = e._R(4);
						break;
					case 5:
						e.l += 4;
						g = e._R(8, "f");
						break;
					case 11:
						e.l += 4;
						g = us(e, 4);
						break;
					case 64:
						e.l += 4;
						g = wr(Wi(e));
						break;
					default:
						throw new Error("unparsed value: " + e[e.l]);
					}
					d[b] = g
				}
			}
		}
		e.l = t + a;
		return d
	}
	var as = ["CodePage", "Thumbnail", "_PID_LINKBASE", "_PID_HLINKS", "SystemIdentifier", "FMTID"];
	function ns(e) {
		switch (typeof e) {
		case "boolean":
			return 11;
		case "number":
			return (e | 0) == e ? 3 : 5;
		case "string":
			return 31;
		case "object":
			if (e instanceof Date) return 64;
			break;
		}
		return - 1
	}
	function is(e, r, t) {
		var a = Ea(8),
		n = [],
		i = [];
		var s = 8,
		f = 0;
		var l = Ea(8),
		o = Ea(8);
		l._W(4, 2);
		l._W(4, 1200);
		o._W(4, 1);
		i.push(l);
		n.push(o);
		s += 8 + l.length;
		if (!r) {
			o = Ea(8);
			o._W(4, 0);
			n.unshift(o);
			var c = [Ea(4)];
			c[0]._W(4, e.length);
			for (f = 0; f < e.length; ++f) {
				var u = e[f][0];
				l = Ea(4 + 4 + 2 * (u.length + 1) + (u.length % 2 ? 0 : 2));
				l._W(4, f + 2);
				l._W(4, u.length + 1);
				l._W(0, u, "dbcs");
				while (l.l != l.length) l._W(1, 0);
				c.push(l)
			}
			l = P(c);
			i.unshift(l);
			s += 8 + l.length
		}
		for (f = 0; f < e.length; ++f) {
			if (r && !r[e[f][0]]) continue;
			if (as.indexOf(e[f][0]) > -1 || Ci.indexOf(e[f][0]) > -1) continue;
			if (e[f][1] == null) continue;
			var h = e[f][1],
			d = 0;
			if (r) {
				d = +r[e[f][0]];
				var v = t[d];
				if (v.p == "version" && typeof h == "string") {
					var p = h.split(".");
					h = ( + p[0] << 16) + ( + p[1] || 0)
				}
				l = rs(v.t, h)
			} else {
				var m = ns(h);
				if (m == -1) {
					m = 31;
					h = String(h)
				}
				l = rs(m, h)
			}
			i.push(l);
			o = Ea(8);
			o._W(4, !r ? 2 + f: d);
			n.push(o);
			s += 8 + l.length
		}
		var b = 8 * (i.length + 1);
		for (f = 0; f < i.length; ++f) {
			n[f]._W(4, b);
			b += i[f].length
		}
		a._W(4, s);
		a._W(4, i.length);
		return P([a].concat(n).concat(i))
	}
	function ss(e, r, t) {
		var a = e.content;
		if (!a) return {};
		Ta(a, 0);
		var n, i, s, f, l = 0;
		a.chk("feff", "Byte Order: ");
		a._R(2);
		var o = a._R(4);
		var c = a._R(16);
		if (c !== Qe.utils.consts.HEADER_CLSID && c !== t) throw new Error("Bad PropertySet CLSID " + c);
		n = a._R(4);
		if (n !== 1 && n !== 2) throw new Error("Unrecognized #Sets: " + n);
		i = a._R(16);
		f = a._R(4);
		if (n === 1 && f !== a.l) throw new Error("Length mismatch: " + f + " !== " + a.l);
		else if (n === 2) {
			s = a._R(16);
			l = a._R(4)
		}
		var u = ts(a, r);
		var h = {
			SystemIdentifier: o
		};
		for (var d in u) h[d] = u[d];
		h.FMTID = i;
		if (n === 1) return h;
		if (l - a.l == 2) a.l += 2;
		if (a.l !== l) throw new Error("Length mismatch 2: " + a.l + " !== " + l);
		var v;
		try {
			v = ts(a, null)
		} catch(p) {}
		for (d in v) h[d] = v[d];
		h.FMTID = [i, s];
		return h
	}
	function fs(e, r, t, a, n, i) {
		var s = Ea(n ? 68 : 48);
		var f = [s];
		s._W(2, 65534);
		s._W(2, 0);
		s._W(4, 842412599);
		s._W(16, Qe.utils.consts.HEADER_CLSID, "hex");
		s._W(4, n ? 2 : 1);
		s._W(16, r, "hex");
		s._W(4, n ? 68 : 48);
		var l = is(e, t, a);
		f.push(l);
		if (n) {
			var o = is(n, null, null);
			s._W(16, i, "hex");
			s._W(4, 68 + l.length);
			f.push(o)
		}
		return P(f)
	}
	function ls(e, r) {
		e._R(r);
		return null
	}
	function os(e, r) {
		if (!r) r = Ea(e);
		for (var t = 0; t < e; ++t) r._W(1, 0);
		return r
	}
	function cs(e, r, t) {
		var a = [],
		n = e.l + r;
		while (e.l < n) a.push(t(e, n - e.l));
		if (n !== e.l) throw new Error("Slurp error");
		return a
	}
	function us(e, r) {
		return e._R(r) === 1
	}
	function hs(e, r) {
		if (!r) r = Ea(2);
		r._W(2, +!!e);
		return r
	}
	function ds(e) {
		return e._R(2, "u")
	}
	function vs(e, r) {
		if (!r) r = Ea(2);
		r._W(2, e);
		return r
	}
	function ps(e, r) {
		return cs(e, r, ds)
	}
	function ms(e) {
		var r = e._R(1),
		t = e._R(1);
		return t === 1 ? r: r === 1
	}
	function bs(e, r, t) {
		if (!t) t = Ea(2);
		t._W(1, r == "e" ? +e: +!!e);
		t._W(1, r == "e" ? 1 : 0);
		return t
	}
	function gs(e, t, a) {
		var n = e._R(a && a.biff >= 12 ? 2 : 1);
		var i = "sbcs-cont";
		var s = r;
		if (a && a.biff >= 8) r = 1200;
		if (!a || a.biff == 8) {
			var f = e._R(1);
			if (f) {
				i = "dbcs-cont"
			}
		} else if (a.biff == 12) {
			i = "wstr"
		}
		if (a.biff >= 2 && a.biff <= 5) i = "cpstr";
		var l = n ? e._R(n, i) : "";
		r = s;
		return l
	}
	function ws(e) {
		var t = r;
		r = 1200;
		var a = e._R(2),
		n = e._R(1);
		var i = n & 4,
		s = n & 8;
		var f = 1 + (n & 1);
		var l = 0,
		o;
		var c = {};
		if (s) l = e._R(2);
		if (i) o = e._R(4);
		var u = f == 2 ? "dbcs-cont": "sbcs-cont";
		var h = a === 0 ? "": e._R(a, u);
		if (s) e.l += 4 * l;
		if (i) e.l += o;
		c.t = h;
		if (!s) {
			c.raw = "<t>" + c.t + "</t>";
			c.r = c.t
		}
		r = t;
		return c
	}
	function ks(e) {
		var r = e.t || "",
		t = 1;
		var a = Ea(3 + (t > 1 ? 2 : 0));
		a._W(2, r.length);
		a._W(1, (t > 1 ? 8 : 0) | 1);
		if (t > 1) a._W(2, t);
		var n = Ea(2 * r.length);
		n._W(2 * r.length, r, "utf16le");
		var i = [a, n];
		return P(i)
	}
	function Ts(e, r, t) {
		var a;
		if (t) {
			if (t.biff >= 2 && t.biff <= 5) return e._R(r, "cpstr");
			if (t.biff >= 12) return e._R(r, "dbcs-cont")
		}
		var n = e._R(1);
		if (n === 0) {
			a = e._R(r, "sbcs-cont")
		} else {
			a = e._R(r, "dbcs-cont")
		}
		return a
	}
	function ys(e, r, t) {
		var a = e._R(t && t.biff == 2 ? 1 : 2);
		if (a === 0) {
			e.l++;
			return ""
		}
		return Ts(e, a, t)
	}
	function Es(e, r, t) {
		if (t.biff > 5) return ys(e, r, t);
		var a = e._R(1);
		if (a === 0) {
			e.l++;
			return ""
		}
		return e._R(a, t.biff <= 4 || !e.lens ? "cpstr": "sbcs-cont")
	}
	function _s(e, r, t) {
		if (!t) t = Ea(3 + 2 * e.length);
		t._W(2, e.length);
		t._W(1, 1);
		t._W(31, e, "utf16le");
		return t
	}
	function Ss(e) {
		var r = e._R(1);
		e.l++;
		var t = e._R(2);
		e.l += 2;
		return [r, t]
	}
	function xs(e) {
		var r = e._R(4),
		t = e.l;
		var a = false;
		if (r > 24) {
			e.l += r - 24;
			if (e._R(16) === "795881f43b1d7f48af2c825dc4852763") a = true;
			e.l = t
		}
		var n = e._R((a ? r - 24 : r) >> 1, "utf16le").replace(M, "");
		if (a) e.l += 24;
		return n
	}
	function As(e) {
		var r = e._R(2);
		var t = "";
		while (r-->0) t += "../";
		var a = e._R(0, "lpstr-ansi");
		e.l += 2;
		if (e._R(2) != 57005) throw new Error("Bad FileMoniker");
		var n = e._R(4);
		if (n === 0) return t + a.replace(/\\/g, "/");
		var i = e._R(4);
		if (e._R(2) != 3) throw new Error("Bad FileMoniker");
		var s = e._R(i >> 1, "utf16le").replace(M, "");
		return t + s
	}
	function Cs(e, r) {
		var t = e._R(16);
		r -= 16;
		switch (t) {
		case "e0c9ea79f9bace118c8200aa004ba90b":
			return xs(e, r);
		case "0303000000000000c000000000000046":
			return As(e, r);
		default:
			throw new Error("Unsupported Moniker " + t);
		}
	}
	function Rs(e) {
		var r = e._R(4);
		var t = r > 0 ? e._R(r, "utf16le").replace(M, "") : "";
		return t
	}
	function Os(e, r) {
		if (!r) r = Ea(6 + e.length * 2);
		r._W(4, 1 + e.length);
		for (var t = 0; t < e.length; ++t) r._W(2, e.charCodeAt(t));
		r._W(2, 0);
		return r
	}
	function Is(e, r) {
		var t = e.l + r;
		var a = e._R(4);
		if (a !== 2) throw new Error("Unrecognized streamVersion: " + a);
		var n = e._R(2);
		e.l += 2;
		var i, s, f, l, o = "",
		c, u;
		if (n & 16) i = Rs(e, t - e.l);
		if (n & 128) s = Rs(e, t - e.l);
		if ((n & 257) === 257) f = Rs(e, t - e.l);
		if ((n & 257) === 1) l = Cs(e, t - e.l);
		if (n & 8) o = Rs(e, t - e.l);
		if (n & 32) c = e._R(16);
		if (n & 64) u = Wi(e);
		e.l = t;
		var h = s || f || l || "";
		if (h && o) h += "#" + o;
		if (!h) h = "#" + o;
		if (n & 2 && h.charAt(0) == "/" && h.charAt(1) != "/") h = "file://" + h;
		var d = {
			Target: h
		};
		if (c) d.guid = c;
		if (u) d.time = u;
		if (i) d.Tooltip = i;
		return d
	}
	function Ns(e) {
		var r = Ea(512),
		t = 0;
		var a = e.Target;
		if (a.slice(0, 7) == "file://") a = a.slice(7);
		var n = a.indexOf("#");
		var i = n > -1 ? 31 : 23;
		switch (a.charAt(0)) {
		case "#":
			i = 28;
			break;
		case ".":
			i &= ~2;
			break;
		}
		r._W(4, 2);
		r._W(4, i);
		var s = [8, 6815827, 6619237, 4849780, 83];
		for (t = 0; t < s.length; ++t) r._W(4, s[t]);
		if (i == 28) {
			a = a.slice(1);
			Os(a, r)
		} else if (i & 2) {
			s = "e0 c9 ea 79 f9 ba ce 11 8c 82 00 aa 00 4b a9 0b".split(" ");
			for (t = 0; t < s.length; ++t) r._W(1, parseInt(s[t], 16));
			var f = n > -1 ? a.slice(0, n) : a;
			r._W(4, 2 * (f.length + 1));
			for (t = 0; t < f.length; ++t) r._W(2, f.charCodeAt(t));
			r._W(2, 0);
			if (i & 8) Os(n > -1 ? a.slice(n + 1) : "", r)
		} else {
			s = "03 03 00 00 00 00 00 00 c0 00 00 00 00 00 00 46".split(" ");
			for (t = 0; t < s.length; ++t) r._W(1, parseInt(s[t], 16));
			var l = 0;
			while (a.slice(l * 3, l * 3 + 3) == "../" || a.slice(l * 3, l * 3 + 3) == "..\\")++l;
			r._W(2, l);
			r._W(4, a.length - 3 * l + 1);
			for (t = 0; t < a.length - 3 * l; ++t) r._W(1, a.charCodeAt(t + 3 * l) & 255);
			r._W(1, 0);
			r._W(2, 65535);
			r._W(2, 57005);
			for (t = 0; t < 6; ++t) r._W(4, 0)
		}
		return r.slice(0, r.l)
	}
	function Fs(e) {
		var r = e._R(1),
		t = e._R(1),
		a = e._R(1),
		n = e._R(1);
		return [r, t, a, n]
	}
	function Ds(e, r) {
		var t = Fs(e, r);
		t[3] = 0;
		return t
	}
	function Ps(e, r, t) {
		var a = e._R(2);
		var n = e._R(2);
		var i = {
			r: a,
			c: n,
			ixfe: 0
		};
		if (t && t.biff == 2 || r == 7) {
			var s = e._R(1);
			i.ixfe = s & 63;
			e.l += 2
		} else i.ixfe = e._R(2);
		return i
	}
	function Ls(e, r, t, a) {
		if (!a) a = Ea(6);
		a._W(2, e);
		a._W(2, r);
		a._W(2, t || 0);
		return a
	}
	function Ms(e) {
		var r = e._R(2);
		var t = e._R(2);
		e.l += 8;
		return {
			type: r,
			flags: t
		}
	}
	function Us(e, r, t) {
		return r === 0 ? "": Es(e, r, t)
	}
	function Bs(e, r, t) {
		var a = t.biff > 8 ? 4 : 2;
		var n = e._R(a),
		i = e._R(a, "i"),
		s = e._R(a, "i");
		return [n, i, s]
	}
	function Ws(e) {
		var r = e._R(2);
		var t = Tn(e);
		return [r, t]
	}
	function zs(e, r, t) {
		e.l += 4;
		r -= 4;
		var a = e.l + r;
		var n = gs(e, r, t);
		var i = e._R(2);
		a -= e.l;
		if (i !== a) throw new Error("Malformed AddinUdf: padding = " + a + " != " + i);
		e.l += i;
		return n
	}
	function Hs(e) {
		var r = e._R(2);
		var t = e._R(2);
		var a = e._R(2);
		var n = e._R(2);
		return {
			s: {
				c: a,
				r: r
			},
			e: {
				c: n,
				r: t
			}
		}
	}
	function Vs(e, r) {
		if (!r) r = Ea(8);
		r._W(2, e.s.r);
		r._W(2, e.e.r);
		r._W(2, e.s.c);
		r._W(2, e.e.c);
		return r
	}
	function $s(e) {
		var r = e._R(2);
		var t = e._R(2);
		var a = e._R(1);
		var n = e._R(1);
		return {
			s: {
				c: a,
				r: r
			},
			e: {
				c: n,
				r: t
			}
		}
	}
	var Xs = $s;
	function Gs(e) {
		e.l += 4;
		var r = e._R(2);
		var t = e._R(2);
		var a = e._R(2);
		e.l += 12;
		return [t, r, a]
	}
	function js(e) {
		var r = {};
		e.l += 4;
		e.l += 16;
		r.fSharedNote = e._R(2);
		e.l += 4;
		return r
	}
	function Ks(e) {
		var r = {};
		e.l += 4;
		e.cf = e._R(2);
		return r
	}
	function Ys(e) {
		e.l += 2;
		e.l += e._R(2)
	}
	var Zs = {
		0 : Ys,
		4 : Ys,
		5 : Ys,
		6 : Ys,
		7 : Ks,
		8 : Ys,
		9 : Ys,
		10 : Ys,
		11 : Ys,
		12 : Ys,
		13 : js,
		14 : Ys,
		15 : Ys,
		16 : Ys,
		17 : Ys,
		18 : Ys,
		19 : Ys,
		20 : Ys,
		21 : Gs
	};
	function Js(e, r) {
		var t = e.l + r;
		var a = [];
		while (e.l < t) {
			var n = e._R(2);
			e.l -= 2;
			try {
				a[n] = Zs[n](e, t - e.l)
			} catch(i) {
				e.l = t;
				return a
			}
		}
		if (e.l != t) e.l = t;
		return a
	}
	function qs(e, r) {
		var t = {
			BIFFVer: 0,
			dt: 0
		};
		t.BIFFVer = e._R(2);
		r -= 2;
		if (r >= 2) {
			t.dt = e._R(2);
			e.l -= 2
		}
		switch (t.BIFFVer) {
		case 1536:
			;
		case 1280:
			;
		case 1024:
			;
		case 768:
			;
		case 512:
			;
		case 2:
			;
		case 7:
			break;
		default:
			if (r > 6) throw new Error("Unexpected BIFF Ver " + t.BIFFVer);
		}
		e._R(r);
		return t
	}
	function Qs(e, r, t) {
		var a = 1536,
		n = 16;
		switch (t.bookType) {
		case "biff8":
			break;
		case "biff5":
			a = 1280;
			n = 8;
			break;
		case "biff4":
			a = 4;
			n = 6;
			break;
		case "biff3":
			a = 3;
			n = 6;
			break;
		case "biff2":
			a = 2;
			n = 4;
			break;
		case "xla":
			break;
		default:
			throw new Error("unsupported BIFF version");
		}
		var i = Ea(n);
		i._W(2, a);
		i._W(2, r);
		if (n > 4) i._W(2, 29282);
		if (n > 6) i._W(2, 1997);
		if (n > 8) {
			i._W(2, 49161);
			i._W(2, 1);
			i._W(2, 1798);
			i._W(2, 0)
		}
		return i
	}
	function ef(e, r) {
		if (r === 0) return 1200;
		if (e._R(2) !== 1200) {}
		return 1200
	}
	function rf(e, r, t) {
		if (t.enc) {
			e.l += r;
			return ""
		}
		var a = e.l;
		var n = Es(e, 0, t);
		e._R(r + a - e.l);
		return n
	}
	function tf(e, r) {
		var t = !r || r.biff == 8;
		var a = Ea(t ? 112 : 54);
		a._W(r.biff == 8 ? 2 : 1, 7);
		if (t) a._W(1, 0);
		a._W(4, 859007059);
		a._W(4, 5458548 | (t ? 0 : 536870912));
		while (a.l < a.length) a._W(1, t ? 0 : 32);
		return a
	}
	function af(e, r, t) {
		var a = t && t.biff == 8 || r == 2 ? e._R(2) : (e.l += r, 0);
		return {
			fDialog: a & 16,
			fBelow: a & 64,
			fRight: a & 128
		}
	}
	function nf(e, r, t) {
		var a = "";
		if (t.biff == 4) {
			a = gs(e, 0, t);
			if (a.length === 0) a = "Sheet1";
			return {
				name: a
			}
		}
		var n = e._R(4);
		var i = e._R(1) & 3;
		var s = e._R(1);
		switch (s) {
		case 0:
			s = "Worksheet";
			break;
		case 1:
			s = "Macrosheet";
			break;
		case 2:
			s = "Chartsheet";
			break;
		case 6:
			s = "VBAModule";
			break;
		}
		a = gs(e, 0, t);
		if (a.length === 0) a = "Sheet1";
		return {
			pos: n,
			hs: i,
			dt: s,
			name: a
		}
	}
	function sf(e, r) {
		var t = !r || r.biff >= 8 ? 2 : 1;
		var a = Ea(8 + t * e.name.length);
		a._W(4, e.pos);
		a._W(1, e.hs || 0);
		a._W(1, e.dt);
		a._W(1, e.name.length);
		if (r.biff >= 8) a._W(1, 1);
		a._W(t * e.name.length, e.name, r.biff < 8 ? "sbcs": "utf16le");
		var n = a.slice(0, a.l);
		n.l = a.l;
		return n
	}
	function ff(e, r) {
		var t = e.l + r;
		var a = e._R(4);
		var n = e._R(4);
		var i = [];
		for (var s = 0; s != n && e.l < t; ++s) {
			i.push(ws(e))
		}
		i.Count = a;
		i.Unique = n;
		return i
	}
	function lf(e, r) {
		var t = Ea(8);
		t._W(4, e.Count);
		t._W(4, e.Unique);
		var a = [];
		for (var n = 0; n < e.length; ++n) a[n] = ks(e[n], r);
		var i = P([t].concat(a));
		i.parts = [t.length].concat(a.map(function(e) {
			return e.length
		}));
		return i
	}
	function of(e, r) {
		var t = {};
		t.dsst = e._R(2);
		e.l += r - 2;
		return t
	}
	function cf(e) {
		var r = {};
		r.r = e._R(2);
		r.c = e._R(2);
		r.cnt = e._R(2) - r.c;
		var t = e._R(2);
		e.l += 4;
		var a = e._R(1);
		e.l += 3;
		if (a & 7) r.level = a & 7;
		if (a & 32) r.hidden = true;
		if (a & 64) r.hpt = t / 20;
		return r
	}
	function uf(e) {
		var r = Ms(e);
		if (r.type != 2211) throw new Error("Invalid Future Record " + r.type);
		var t = e._R(4);
		return t !== 0
	}
	function hf(e) {
		e._R(2);
		return e._R(4)
	}
	function df(e, r, t) {
		var a = 0;
		if (! (t && t.biff == 2)) {
			a = e._R(2)
		}
		var n = e._R(2);
		if (t && t.biff == 2) {
			a = 1 - (n >> 15);
			n &= 32767
		}
		var i = {
			Unsynced: a & 1,
			DyZero: (a & 2) >> 1,
			ExAsc: (a & 4) >> 2,
			ExDsc: (a & 8) >> 3
		};
		return [i, n]
	}
	function vf(e) {
		var r = e._R(2),
		t = e._R(2),
		a = e._R(2),
		n = e._R(2);
		var i = e._R(2),
		s = e._R(2),
		f = e._R(2);
		var l = e._R(2),
		o = e._R(2);
		return {
			Pos: [r, t],
			Dim: [a, n],
			Flags: i,
			CurTab: s,
			FirstTab: f,
			Selected: l,
			TabRatio: o
		}
	}
	function pf() {
		var e = Ea(18);
		e._W(2, 0);
		e._W(2, 0);
		e._W(2, 29280);
		e._W(2, 17600);
		e._W(2, 56);
		e._W(2, 0);
		e._W(2, 0);
		e._W(2, 1);
		e._W(2, 500);
		return e
	}
	function mf(e, r, t) {
		if (t && t.biff >= 2 && t.biff < 5) return {};
		var a = e._R(2);
		return {
			RTL: a & 64
		}
	}
	function bf(e) {
		var r = Ea(18),
		t = 1718;
		if (e && e.RTL) t |= 64;
		r._W(2, t);
		r._W(4, 0);
		r._W(4, 64);
		r._W(4, 0);
		r._W(4, 0);
		return r
	}
	function gf() {}
	function wf(e, r, t) {
		var a = {
			dyHeight: e._R(2),
			fl: e._R(2)
		};
		switch (t && t.biff || 8) {
		case 2:
			break;
		case 3:
			;
		case 4:
			e.l += 2;
			break;
		default:
			e.l += 10;
			break;
		}
		a.name = gs(e, 0, t);
		return a
	}
	function kf(e, r) {
		var t = e.name || "Arial";
		var a = r && r.biff == 5,
		n = a ? 15 + t.length: 16 + 2 * t.length;
		var i = Ea(n);
		i._W(2, (e.sz || 12) * 20);
		i._W(4, 0);
		i._W(2, 400);
		i._W(4, 0);
		i._W(2, 0);
		i._W(1, t.length);
		if (!a) i._W(1, 1);
		i._W((a ? 1 : 2) * t.length, t, a ? "sbcs": "utf16le");
		return i
	}
	function Tf(e, r, t) {
		var a = Ps(e, r, t);
		a.isst = e._R(4);
		return a
	}
	function yf(e, r, t, a) {
		var n = Ea(10);
		Ls(e, r, a, n);
		n._W(4, t);
		return n
	}
	function Ef(e, r, t) {
		if (t.biffguess && t.biff == 2) t.biff = 5;
		var a = e.l + r;
		var n = Ps(e, r, t);
		var i = ys(e, a - e.l, t);
		n.val = i;
		return n
	}
	function _f(e, r, t, a, n) {
		var i = !n || n.biff == 8;
		var s = Ea(6 + 2 + +i + (1 + i) * t.length);
		Ls(e, r, a, s);
		s._W(2, t.length);
		if (i) s._W(1, 1);
		s._W((1 + i) * t.length, t, i ? "utf16le": "sbcs");
		return s
	}
	function Sf(e, r, t) {
		var a = e._R(2);
		var n = Es(e, 0, t);
		return [a, n]
	}
	function xf(e, r, t, a) {
		var n = t && t.biff == 5;
		if (!a) a = Ea(n ? 3 + r.length: 5 + 2 * r.length);
		a._W(2, e);
		a._W(n ? 1 : 2, r.length);
		if (!n) a._W(1, 1);
		a._W((n ? 1 : 2) * r.length, r, n ? "sbcs": "utf16le");
		var i = a.length > a.l ? a.slice(0, a.l) : a;
		if (i.l == null) i.l = i.length;
		return i
	}
	var Af = Es;
	function Cf(e) {
		var r = Ea(1 + e.length);
		r._W(1, e.length);
		r._W(e.length, e, "sbcs");
		return r
	}
	function Rf(e) {
		var r = Ea(3 + e.length);
		r.l += 2;
		r._W(1, e.length);
		r._W(e.length, e, "sbcs");
		return r
	}
	function Of(e, r, t) {
		var a = e.l + r;
		var n = t.biff == 8 || !t.biff ? 4 : 2;
		var i = e._R(n),
		s = e._R(n);
		var f = e._R(2),
		l = e._R(2);
		e.l = a;
		return {
			s: {
				r: i,
				c: f
			},
			e: {
				r: s,
				c: l
			}
		}
	}
	function If(e, r) {
		var t = r.biff == 8 || !r.biff ? 4 : 2;
		var a = Ea(2 * t + 6);
		a._W(t, e.s.r);
		a._W(t, e.e.r + 1);
		a._W(2, e.s.c);
		a._W(2, e.e.c + 1);
		a._W(2, 0);
		return a
	}
	function Nf(e) {
		var r = e._R(2),
		t = e._R(2);
		var a = Ws(e);
		return {
			r: r,
			c: t,
			ixfe: a[0],
			rknum: a[1]
		}
	}
	function Ff(e, r) {
		var t = e.l + r - 2;
		var a = e._R(2),
		n = e._R(2);
		var i = [];
		while (e.l < t) i.push(Ws(e));
		if (e.l !== t) throw new Error("MulRK read error");
		var s = e._R(2);
		if (i.length != s - n + 1) throw new Error("MulRK length mismatch");
		return {
			r: a,
			c: n,
			C: s,
			rkrec: i
		}
	}
	function Df(e, r) {
		var t = e.l + r - 2;
		var a = e._R(2),
		n = e._R(2);
		var i = [];
		while (e.l < t) i.push(e._R(2));
		if (e.l !== t) throw new Error("MulBlank read error");
		var s = e._R(2);
		if (i.length != s - n + 1) throw new Error("MulBlank length mismatch");
		return {
			r: a,
			c: n,
			C: s,
			ixfe: i
		}
	}
	function Pf(e, r, t, a) {
		var n = {};
		var i = e._R(4),
		s = e._R(4);
		var f = e._R(4),
		l = e._R(2);
		n.patternType = qn[f >> 26];
		if (!a.cellStyles) return n;
		n.alc = i & 7;
		n.fWrap = i >> 3 & 1;
		n.alcV = i >> 4 & 7;
		n.fJustLast = i >> 7 & 1;
		n.trot = i >> 8 & 255;
		n.cIndent = i >> 16 & 15;
		n.fShrinkToFit = i >> 20 & 1;
		n.iReadOrder = i >> 22 & 2;
		n.fAtrNum = i >> 26 & 1;
		n.fAtrFnt = i >> 27 & 1;
		n.fAtrAlc = i >> 28 & 1;
		n.fAtrBdr = i >> 29 & 1;
		n.fAtrPat = i >> 30 & 1;
		n.fAtrProt = i >> 31 & 1;
		n.dgLeft = s & 15;
		n.dgRight = s >> 4 & 15;
		n.dgTop = s >> 8 & 15;
		n.dgBottom = s >> 12 & 15;
		n.icvLeft = s >> 16 & 127;
		n.icvRight = s >> 23 & 127;
		n.grbitDiag = s >> 30 & 3;
		n.icvTop = f & 127;
		n.icvBottom = f >> 7 & 127;
		n.icvDiag = f >> 14 & 127;
		n.dgDiag = f >> 21 & 15;
		n.icvFore = l & 127;
		n.icvBack = l >> 7 & 127;
		n.fsxButton = l >> 14 & 1;
		return n
	}
	function Lf(e, r, t) {
		var a = {};
		a.ifnt = e._R(2);
		a.numFmtId = e._R(2);
		a.flags = e._R(2);
		a.fStyle = a.flags >> 2 & 1;
		r -= 6;
		a.data = Pf(e, r, a.fStyle, t);
		return a
	}
	function Mf(e, r, t, a) {
		var n = t && t.biff == 5;
		if (!a) a = Ea(n ? 16 : 20);
		a._W(2, 0);
		if (e.style) {
			a._W(2, e.numFmtId || 0);
			a._W(2, 65524)
		} else {
			a._W(2, e.numFmtId || 0);
			a._W(2, r << 4)
		}
		var i = 0;
		if (e.numFmtId > 0 && n) i |= 1024;
		a._W(4, i);
		a._W(4, 0);
		if (!n) a._W(4, 0);
		a._W(2, 0);
		return a
	}
	function Uf(e) {
		var r = {};
		r.ifnt = e._R(1);
		e.l++;
		r.flags = e._R(1);
		r.numFmtId = r.flags & 63;
		r.flags >>= 6;
		r.fStyle = 0;
		r.data = {};
		return r
	}
	function Bf(e) {
		var r = Ea(4);
		r.l += 2;
		r._W(1, e.numFmtId);
		r.l++;
		return r
	}
	function Wf(e) {
		var r = Ea(12);
		r.l++;
		r._W(1, e.numFmtId);
		r.l += 10;
		return r
	}
	var zf = Wf;
	function Hf(e) {
		var r = {};
		r.ifnt = e._R(1);
		r.numFmtId = e._R(1);
		r.flags = e._R(2);
		r.fStyle = r.flags >> 2 & 1;
		r.data = {};
		return r
	}
	function Vf(e) {
		var r = {};
		r.ifnt = e._R(1);
		r.numFmtId = e._R(1);
		r.flags = e._R(2);
		r.fStyle = r.flags >> 2 & 1;
		r.data = {};
		return r
	}
	function $f(e) {
		e.l += 4;
		var r = [e._R(2), e._R(2)];
		if (r[0] !== 0) r[0]--;
		if (r[1] !== 0) r[1]--;
		if (r[0] > 7 || r[1] > 7) throw new Error("Bad Gutters: " + r.join("|"));
		return r
	}
	function Xf(e) {
		var r = Ea(8);
		r._W(4, 0);
		r._W(2, e[0] ? e[0] + 1 : 0);
		r._W(2, e[1] ? e[1] + 1 : 0);
		return r
	}
	function Gf(e, r, t) {
		var a = Ps(e, 6, t);
		var n = ms(e, 2);
		a.val = n;
		a.t = n === true || n === false ? "b": "e";
		return a
	}
	function jf(e, r, t, a, n, i) {
		var s = Ea(8);
		Ls(e, r, a, s);
		bs(t, i, s);
		return s
	}
	function Kf(e, r, t) {
		if (t.biffguess && t.biff == 2) t.biff = 5;
		var a = Ps(e, 6, t);
		var n = An(e, 8);
		a.val = n;
		return a
	}
	function Yf(e, r, t, a) {
		var n = Ea(14);
		Ls(e, r, a, n);
		Cn(t, n);
		return n
	}
	var Zf = Us;
	function Jf(e, r, t) {
		var a = e.l + r;
		var n = e._R(2);
		var i = e._R(2);
		t.sbcch = i;
		if (i == 1025 || i == 14849) return [i, n];
		if (i < 1 || i > 255) throw new Error("Unexpected SupBook type: " + i);
		var s = Ts(e, i);
		var f = [];
		while (a > e.l) f.push(ys(e));
		return [i, n, s, f]
	}
	function qf(e, r, t) {
		var a = e._R(2);
		var n;
		var i = {
			fBuiltIn: a & 1,
			fWantAdvise: a >>> 1 & 1,
			fWantPict: a >>> 2 & 1,
			fOle: a >>> 3 & 1,
			fOleLink: a >>> 4 & 1,
			cf: a >>> 5 & 1023,
			fIcon: a >>> 15 & 1
		};
		if (t.sbcch === 14849) n = zs(e, r - 2, t);
		i.body = n || e._R(r - 2);
		if (typeof n === "string") i.Name = n;
		return i
	}
	function Qf(e, r, t) {
		var a = e.l + r;
		var n = e._R(2);
		var i = e._R(1);
		var s = e._R(1);
		var f = e._R(t && t.biff == 2 ? 1 : 2);
		var l = 0;
		if (!t || t.biff >= 5) {
			if (t.biff != 5) e.l += 2;
			l = e._R(2);
			if (t.biff == 5) e.l += 2;
			e.l += 4
		}
		var o = Ts(e, s, t);
		if (n & 32) o = ni[o.charCodeAt(0)];
		var c = a - e.l;
		if (t && t.biff == 2)--c;
		var u = a == e.l || f === 0 || !(c > 0) ? [] : Wd(e, c, t, f);
		return {
			chKey: i,
			Name: o,
			itab: l,
			rgce: u
		}
	}
	function el(e, r, t) {
		if (t.biff < 8) return rl(e, r, t);
		if (! (t.biff > 8) && r == e[e.l] + (e[e.l + 1] == 3 ? 1 : 0) + 1) return rl(e, r, t);
		var a = [],
		n = e.l + r,
		i = e._R(t.biff > 8 ? 4 : 2);
		while (i--!==0) a.push(Bs(e, t.biff > 8 ? 12 : 6, t));
		if (e.l != n) throw new Error("Bad ExternSheet: " + e.l + " != " + n);
		return a
	}
	function rl(e, r, t) {
		if (e[e.l + 1] == 3) e[e.l]++;
		var a = gs(e, r, t);
		return a.charCodeAt(0) == 3 ? a.slice(1) : a
	}
	function tl(e, r, t) {
		if (t.biff < 8) {
			e.l += r;
			return
		}
		var a = e._R(2);
		var n = e._R(2);
		var i = Ts(e, a, t);
		var s = Ts(e, n, t);
		return [i, s]
	}
	function al(e, r, t) {
		var a = $s(e, 6);
		e.l++;
		var n = e._R(1);
		r -= 8;
		return [zd(e, r, t), n, a]
	}
	function nl(e, r, t) {
		var a = Xs(e, 6);
		switch (t.biff) {
		case 2:
			e.l++;
			r -= 7;
			break;
		case 3:
			;
		case 4:
			e.l += 2;
			r -= 8;
			break;
		default:
			e.l += 6;
			r -= 12;
		}
		return [a, Ud(e, r, t, a)]
	}
	function il(e) {
		var r = e._R(4) !== 0;
		var t = e._R(4) !== 0;
		var a = e._R(4);
		return [r, t, a]
	}
	function sl(e, r, t) {
		var a = e._R(2),
		n = e._R(2);
		var i = e._R(2),
		s = e._R(2);
		var f = Es(e, 0, t);
		return [{
			r: a,
			c: n
		},
		f, s, i]
	}
	function fl(e, r, t) {
		if (t && t.biff < 8) {
			var a = e._R(2),
			n = e._R(2);
			if (a == 65535 || a == -1) return;
			var i = e._R(2);
			var s = e._R(Math.min(i, 2048), "cpstr");
			return [{
				r: a,
				c: n
			},
			s]
		}
		return sl(e, r, t)
	}
	function ll(e, r, t, a) {
		var n = Ea(6 + (a || e.length));
		n._W(2, r);
		n._W(2, t);
		n._W(2, a || e.length);
		n._W(e.length, e, "sbcs");
		return n
	}
	function ol(e, r) {
		var t = [];
		var a = e._R(2);
		while (a--) t.push(Hs(e, r));
		return t
	}
	function cl(e) {
		var r = Ea(2 + e.length * 8);
		r._W(2, e.length);
		for (var t = 0; t < e.length; ++t) Vs(e[t], r);
		return r
	}
	function ul(e, r, t) {
		if (t && t.biff < 8) return dl(e, r, t);
		var a = Gs(e, 22);
		var n = Js(e, r - 22, a[1]);
		return {
			cmo: a,
			ft: n
		}
	}
	var hl = {
		8 : function(e, r) {
			var t = e.l + r;
			e.l += 10;
			var a = e._R(2);
			e.l += 4;
			e.l += 2;
			e.l += 2;
			e.l += 2;
			e.l += 4;
			var n = e._R(1);
			e.l += n;
			e.l = t;
			return {
				fmt: a
			}
		}
	};
	function dl(e, r, t) {
		e.l += 4;
		var a = e._R(2);
		var n = e._R(2);
		var i = e._R(2);
		e.l += 2;
		e.l += 2;
		e.l += 2;
		e.l += 2;
		e.l += 2;
		e.l += 2;
		e.l += 2;
		e.l += 2;
		e.l += 2;
		e.l += 6;
		r -= 36;
		var s = [];
		s.push((hl[a] || ya)(e, r, t));
		return {
			cmo: [n, a, i],
			ft: s
		}
	}
	function vl(e, r, t) {
		var a = e.l;
		var n = "";
		try {
			e.l += 4;
			var i = (t.lastobj || {
				cmo: [0, 0]
			}).cmo[1];
			var s;
			if ([0, 5, 7, 11, 12, 14].indexOf(i) == -1) e.l += 6;
			else s = Ss(e, 6, t);
			var f = e._R(2);
			e._R(2);
			ds(e, 2);
			var l = e._R(2);
			e.l += l;
			for (var o = 1; o < e.lens.length - 1; ++o) {
				if (e.l - a != e.lens[o]) throw new Error("TxO: bad continue record");
				var c = e[e.l];
				var u = Ts(e, e.lens[o + 1] - e.lens[o] - 1);
				n += u;
				if (n.length >= (c ? f: 2 * f)) break
			}
			if (n.length !== f && n.length !== f * 2) {
				throw new Error("cchText: " + f + " != " + n.length)
			}
			e.l = a + r;
			return {
				t: n
			}
		} catch(h) {
			e.l = a + r;
			return {
				t: n
			}
		}
	}
	function pl(e, r) {
		var t = Hs(e, 8);
		e.l += 16;
		var a = Is(e, r - 24);
		return [t, a]
	}
	function ml(e) {
		var r = Ea(24);
		var t = Wa(e[0]);
		r._W(2, t.r);
		r._W(2, t.r);
		r._W(2, t.c);
		r._W(2, t.c);
		var a = "d0 c9 ea 79 f9 ba ce 11 8c 82 00 aa 00 4b a9 0b".split(" ");
		for (var n = 0; n < 16; ++n) r._W(1, parseInt(a[n], 16));
		return P([r, Ns(e[1])])
	}
	function bl(e, r) {
		e._R(2);
		var t = Hs(e, 8);
		var a = e._R((r - 10) / 2, "dbcs-cont");
		a = a.replace(M, "");
		return [t, a]
	}
	function gl(e) {
		var r = e[1].Tooltip;
		var t = Ea(10 + 2 * (r.length + 1));
		t._W(2, 2048);
		var a = Wa(e[0]);
		t._W(2, a.r);
		t._W(2, a.r);
		t._W(2, a.c);
		t._W(2, a.c);
		for (var n = 0; n < r.length; ++n) t._W(2, r.charCodeAt(n));
		t._W(2, 0);
		return t
	}
	function wl(e) {
		var r = [0, 0],
		t;
		t = e._R(2);
		r[0] = Jn[t] || t;
		t = e._R(2);
		r[1] = Jn[t] || t;
		return r
	}
	function kl(e) {
		if (!e) e = Ea(4);
		e._W(2, 1);
		e._W(2, 1);
		return e
	}
	function Tl(e) {
		var r = e._R(2);
		var t = [];
		while (r-->0) t.push(Ds(e, 8));
		return t
	}
	function yl(e) {
		var r = e._R(2);
		var t = [];
		while (r-->0) t.push(Ds(e, 8));
		return t
	}
	function El(e) {
		e.l += 2;
		var r = {
			cxfs: 0,
			crc: 0
		};
		r.cxfs = e._R(2);
		r.crc = e._R(4);
		return r
	}
	function _l(e, r, t) {
		if (!t.cellStyles) return ya(e, r);
		var a = t && t.biff >= 12 ? 4 : 2;
		var n = e._R(a);
		var i = e._R(a);
		var s = e._R(a);
		var f = e._R(a);
		var l = e._R(2);
		if (a == 2) e.l += 2;
		var o = {
			s: n,
			e: i,
			w: s,
			ixfe: f,
			flags: l
		};
		if (t.biff >= 5 || !t.biff) o.level = l >> 8 & 7;
		return o
	}
	function Sl(e, r) {
		var t = Ea(12);
		t._W(2, r);
		t._W(2, r);
		t._W(2, e.width * 256);
		t._W(2, 0);
		var a = 0;
		if (e.hidden) a |= 1;
		t._W(1, a);
		a = e.level || 0;
		t._W(1, a);
		t._W(2, 0);
		return t
	}
	function xl(e, r) {
		var t = {};
		if (r < 32) return t;
		e.l += 16;
		t.header = An(e, 8);
		t.footer = An(e, 8);
		e.l += 2;
		return t
	}
	function Al(e, r, t) {
		var a = {
			area: false
		};
		if (t.biff != 5) {
			e.l += r;
			return a
		}
		var n = e._R(1);
		e.l += 3;
		if (n & 16) a.area = true;
		return a
	}
	function Cl(e) {
		var r = Ea(2 * e);
		for (var t = 0; t < e; ++t) r._W(2, t + 1);
		return r
	}
	var Rl = Ps;
	var Ol = ps;
	var Il = ys;
	function Nl(e) {
		var r = e._R(2);
		var t = e._R(2);
		var a = e._R(4);
		var n = {
			fmt: r,
			env: t,
			len: a,
			data: e.slice(e.l, e.l + a)
		};
		e.l += a;
		return n
	}
	function Fl(e, r, t, a, n) {
		if (!e) e = Ea(7);
		e._W(2, r);
		e._W(2, t);
		e._W(1, a || 0);
		e._W(1, n || 0);
		e._W(1, 0);
		return e
	}
	function Dl(e, r, t) {
		if (t.biffguess && t.biff == 5) t.biff = 2;
		var a = Ps(e, 7, t);
		var n = Es(e, r - 7, t);
		a.t = "str";
		a.val = n;
		return a
	}
	function Pl(e, r, t) {
		var a = Ps(e, 7, t);
		var n = An(e, 8);
		a.t = "n";
		a.val = n;
		return a
	}
	function Ll(e, r, t, a, n) {
		var i = Ea(15);
		Fl(i, e, r, a || 0, n || 0);
		i._W(8, t, "f");
		return i
	}
	function Ml(e, r, t) {
		var a = Ps(e, 7, t);
		var n = e._R(2);
		a.t = "n";
		a.val = n;
		return a
	}
	function Ul(e, r, t, a, n) {
		var i = Ea(9);
		Fl(i, e, r, a || 0, n || 0);
		i._W(2, t);
		return i
	}
	function Bl(e) {
		var r = e._R(1);
		if (r === 0) {
			e.l++;
			return ""
		}
		return e._R(r, "sbcs-cont")
	}
	function Wl(e, r, t) {
		var a = e.l + 7;
		var n = Ps(e, 6, t);
		e.l = a;
		var i = ms(e, 2);
		n.val = i;
		n.t = i === true || i === false ? "b": "e";
		return n
	}
	function zl(e, r) {
		e.l += 6;
		e.l += 2;
		e.l += 1;
		e.l += 3;
		e.l += 1;
		e.l += r - 13
	}
	function Hl(e, r, t) {
		var a = e.l + r;
		var n = Ps(e, 6, t);
		var i = e._R(2);
		var s = Ts(e, i, t);
		e.l = a;
		n.t = "str";
		n.val = s;
		return n
	}
	function Vl(e) {
		var r = e._R(4);
		var t = e._R(1),
		a = e._R(t, "sbcs");
		if (a.length === 0) a = "Sheet1";
		return {
			flags: r,
			name: a
		}
	}
	var $l = [2, 3, 48, 49, 131, 139, 140, 245];
	var Xl = function() {
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
			255 : 16969
		};
		var n = fr({
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
			0 : 20127
		});
		function i(r, t) {
			var n = [];
			var i = C(1);
			switch (t.type) {
			case "base64":
				i = O(_(r));
				break;
			case "binary":
				i = O(r);
				break;
			case "buffer":
				;
			case "array":
				i = r;
				break;
			}
			Ta(i, 0);
			var s = i._R(1);
			var f = !!(s & 136);
			var l = false,
			o = false;
			switch (s) {
			case 2:
				break;
			case 3:
				break;
			case 48:
				l = true;
				f = true;
				break;
			case 49:
				l = true;
				f = true;
				break;
			case 131:
				break;
			case 139:
				break;
			case 140:
				o = true;
				break;
			case 245:
				break;
			default:
				throw new Error("DBF Unsupported Version: " + s.toString(16));
			}
			var c = 0,
			u = 521;
			if (s == 2) c = i._R(2);
			i.l += 3;
			if (s != 2) c = i._R(4);
			if (c > 1048576) c = 1e6;
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
			if (o) i.l += 36;
			var v = [],
			p = {};
			var m = Math.min(i.length, s == 2 ? 521 : u - 10 - (l ? 264 : 0));
			var b = o ? 32 : 11;
			while (i.l < m && i[i.l] != 13) {
				p = {};
				p.name = (typeof a !== "undefined" ? a.utils.decode(d, i.slice(i.l, i.l + b)) : N(i.slice(i.l, i.l + b))).replace(/[\u0000\r\n].*$/g, "");
				i.l += b;
				p.type = String.fromCharCode(i._R(1));
				if (s != 2 && !o) p.offset = i._R(4);
				p.len = i._R(1);
				if (s == 2) p.offset = i._R(2);
				p.dec = i._R(1);
				if (p.name.length) v.push(p);
				if (s != 2) i.l += o ? 13 : 14;
				switch (p.type) {
				case "B":
					if ((!l || p.len != 8) && t.WTF) console.log("Skipping " + p.name + ":" + p.type);
					break;
				case "G":
					;
				case "P":
					if (t.WTF) console.log("Skipping " + p.name + ":" + p.type);
					break;
				case "+":
					;
				case "0":
					;
				case "@":
					;
				case "C":
					;
				case "D":
					;
				case "F":
					;
				case "I":
					;
				case "L":
					;
				case "M":
					;
				case "N":
					;
				case "O":
					;
				case "T":
					;
				case "Y":
					break;
				default:
					throw new Error("Unknown Field Type: " + p.type);
				}
			}
			if (i[i.l] !== 13) i.l = u - 1;
			if (i._R(1) !== 13) throw new Error("DBF Terminator not found " + i.l + " " + i[i.l]);
			i.l = u;
			var g = 0,
			w = 0;
			n[0] = [];
			for (w = 0; w != v.length; ++w) n[0][w] = v[w].name;
			while (c-->0) {
				if (i[i.l] === 42) {
					i.l += h;
					continue
				}++i.l;
				n[++g] = [];
				w = 0;
				for (w = 0; w != v.length; ++w) {
					var k = i.slice(i.l, i.l + v[w].len);
					i.l += v[w].len;
					Ta(k, 0);
					var T = typeof a !== "undefined" ? a.utils.decode(d, k) : N(k);
					switch (v[w].type) {
					case "C":
						if (T.trim().length) n[g][w] = T.replace(/\s+$/, "");
						break;
					case "D":
						if (T.length === 8) {
							n[g][w] = new Date(Date.UTC( + T.slice(0, 4), +T.slice(4, 6) - 1, +T.slice(6, 8), 0, 0, 0, 0));
							if (! (t && t.UTC)) {
								n[g][w] = Fr(n[g][w])
							}
						} else n[g][w] = T;
						break;
					case "F":
						n[g][w] = parseFloat(T.trim());
						break;
					case "+":
						;
					case "I":
						n[g][w] = o ? k._R(-4, "i") ^ 2147483648 : k._R(4, "i");
						break;
					case "L":
						switch (T.trim().toUpperCase()) {
						case "Y":
							;
						case "T":
							n[g][w] = true;
							break;
						case "N":
							;
						case "F":
							n[g][w] = false;
							break;
						case "":
							;
						case "\0":
							;
						case "?":
							break;
						default:
							throw new Error("DBF Unrecognized L:|" + T + "|");
						}
						break;
					case "M":
						if (!f) throw new Error("DBF Unexpected MEMO for type " + s.toString(16));
						n[g][w] = "##MEMO##" + (o ? parseInt(T.trim(), 10) : k._R(4));
						break;
					case "N":
						T = T.replace(/\u0000/g, "").trim();
						if (T && T != ".") n[g][w] = +T || 0;
						break;
					case "@":
						n[g][w] = new Date(k._R(-8, "f") - 621356832e5);
						break;
					case "T":
						{
							var y = k._R(4),
							E = k._R(4);
							if (y == 0 && E == 0) break;
							n[g][w] = new Date((y - 2440588) * 864e5 + E);
							if (! (t && t.UTC)) n[g][w] = Fr(n[g][w])
						}
						break;
					case "Y":
						n[g][w] = k._R(4, "i") / 1e4 + k._R(4, "i") / 1e4 * Math.pow(2, 32);
						break;
					case "O":
						n[g][w] = -k._R(-8, "f");
						break;
					case "B":
						if (l && v[w].len == 8) {
							n[g][w] = k._R(8, "f");
							break
						};
					case "G":
						;
					case "P":
						k.l += v[w].len;
						break;
					case "0":
						if (v[w].name === "_NullFlags") break;
					default:
						throw new Error("DBF Unsupported data type " + v[w].type);
					}
				}
			}
			if (s != 2) if (i.l < i.length && i[i.l++] != 26) throw new Error("DBF EOF Marker missing " + (i.l - 1) + " of " + i.length + " " + i[i.l - 1].toString(16));
			if (t && t.sheetRows) n = n.slice(0, t.sheetRows);
			t.DBF = v;
			return n
		}
		function s(e, r) {
			var t = r || {};
			if (!t.dateNF) t.dateNF = "yyyymmdd";
			var a = qa(i(e, t), t);
			a["!cols"] = t.DBF.map(function(e) {
				return {
					wch: e.len,
					DBF: e
				}
			});
			delete t.DBF;
			return a
		}
		function f(e, r) {
			try {
				var t = Ya(s(e, r), r);
				t.bookType = "dbf";
				return t
			} catch(a) {
				if (r && r.WTF) throw a
			}
			return {
				SheetNames: [],
				Sheets: {}
			}
		}
		var o = {
			B: 8,
			C: 250,
			L: 1,
			D: 8,
			"?": 0,
			"": 0
		};
		function c(i, s) {
			if (!i["!ref"]) throw new Error("Cannot export empty sheet to DBF");
			var f = s || {};
			var c = r;
			if ( + f.codepage >= 0) l( + f.codepage);
			if (f.type == "string") throw new Error("Cannot write DBF to JS string");
			var u = Sa();
			var h = Mk(i, {
				header: 1,
				raw: true,
				cellDates: true
			});
			var d = h[0],
			v = h.slice(1),
			p = i["!cols"] || [];
			var m = 0,
			b = 0,
			g = 0,
			w = 1;
			for (m = 0; m < d.length; ++m) {
				if (((p[m] || {}).DBF || {}).name) {
					d[m] = p[m].DBF.name; ++g;
					continue
				}
				if (d[m] == null) continue; ++g;
				if (typeof d[m] === "number") d[m] = d[m].toString(10);
				if (typeof d[m] !== "string") throw new Error("DBF Invalid column name " + d[m] + " |" + typeof d[m] + "|");
				if (d.indexOf(d[m]) !== m) for (b = 0; b < 1024; ++b) if (d.indexOf(d[m] + "_" + b) == -1) {
					d[m] += "_" + b;
					break
				}
			}
			var k = Ga(i["!ref"]);
			var T = [];
			var y = [];
			var E = [];
			for (m = 0; m <= k.e.c - k.s.c; ++m) {
				var _ = "",
				S = "",
				x = 0;
				var A = [];
				for (b = 0; b < v.length; ++b) {
					if (v[b][m] != null) A.push(v[b][m])
				}
				if (A.length == 0 || d[m] == null) {
					T[m] = "?";
					continue
				}
				for (b = 0; b < A.length; ++b) {
					switch (typeof A[b]) {
					case "number":
						S = "B";
						break;
					case "string":
						S = "C";
						break;
					case "boolean":
						S = "L";
						break;
					case "object":
						S = A[b] instanceof Date ? "D": "C";
						break;
					default:
						S = "C";
					}
					x = Math.max(x, (typeof a !== "undefined" && typeof A[b] == "string" ? a.utils.encode(t, A[b]) : String(A[b])).length);
					_ = _ && _ != S ? "C": S
				}
				if (x > 250) x = 250;
				S = ((p[m] || {}).DBF || {}).type;
				if (S == "C") {
					if (p[m].DBF.len > x) x = p[m].DBF.len
				}
				if (_ == "B" && S == "N") {
					_ = "N";
					E[m] = p[m].DBF.dec;
					x = p[m].DBF.len
				}
				y[m] = _ == "C" || S == "N" ? x: o[_] || 0;
				w += y[m];
				T[m] = _
			}
			var C = u.next(32);
			C._W(4, 318902576);
			C._W(4, v.length);
			C._W(2, 296 + 32 * g);
			C._W(2, w);
			for (m = 0; m < 4; ++m) C._W(4, 0);
			var R = +n[r] || 3;
			C._W(4, 0 | R << 8);
			if (e[R] != +f.codepage) {
				if (f.codepage) console.error("DBF Unsupported codepage " + r + ", using 1252");
				r = 1252
			}
			for (m = 0, b = 0; m < d.length; ++m) {
				if (d[m] == null) continue;
				var O = u.next(32);
				var I = (d[m].slice(-10) + "\0\0\0\0\0\0\0\0\0\0\0").slice(0, 11);
				O._W(1, I, "sbcs");
				O._W(1, T[m] == "?" ? "C": T[m], "sbcs");
				O._W(4, b);
				O._W(1, y[m] || o[T[m]] || 0);
				O._W(1, E[m] || 0);
				O._W(1, 2);
				O._W(4, 0);
				O._W(1, 0);
				O._W(4, 0);
				O._W(4, 0);
				b += y[m] || o[T[m]] || 0
			}
			var N = u.next(264);
			N._W(4, 13);
			for (m = 0; m < 65; ++m) N._W(4, 0);
			for (m = 0; m < v.length; ++m) {
				var F = u.next(w);
				F._W(1, 0);
				for (b = 0; b < d.length; ++b) {
					if (d[b] == null) continue;
					switch (T[b]) {
					case "L":
						F._W(1, v[m][b] == null ? 63 : v[m][b] ? 84 : 70);
						break;
					case "B":
						F._W(8, v[m][b] || 0, "f");
						break;
					case "N":
						var D = "0";
						if (typeof v[m][b] == "number") D = v[m][b].toFixed(E[b] || 0);
						if (D.length > y[b]) D = D.slice(0, y[b]);
						for (g = 0; g < y[b] - D.length; ++g) F._W(1, 32);
						F._W(1, D, "sbcs");
						break;
					case "D":
						if (!v[m][b]) F._W(8, "00000000", "sbcs");
						else {
							F._W(4, ("0000" + v[m][b].getFullYear()).slice(-4), "sbcs");
							F._W(2, ("00" + (v[m][b].getMonth() + 1)).slice(-2), "sbcs");
							F._W(2, ("00" + v[m][b].getDate()).slice(-2), "sbcs")
						}
						break;
					case "C":
						var P = F.l;
						var L = String(v[m][b] != null ? v[m][b] : "").slice(0, y[b]);
						F._W(1, L, "cpstr");
						P += y[b] - F.l;
						for (g = 0; g < P; ++g) F._W(1, 32);
						break;
					}
				}
			}
			r = c;
			u.next(1)._W(1, 26);
			return u.end()
		}
		return {
			to_workbook: f,
			to_sheet: s,
			from_sheet: c
		}
	} ();
	var Gl = function() {
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
			"{": 223
		};
		var r = new RegExp("N(" + ir(e).join("|").replace(/\|\|\|/, "|\\||").replace(/([?()+])/g, "\\$1").replace("{", "\\{") + "|\\|)", "gm");
		try {
			r = new RegExp("N(" + ir(e).join("|").replace(/\|\|\|/, "|\\||").replace(/([?()+])/g, "\\$1") + "|\\|)", "gm")
		} catch(t) {}
		var n = function(r, t) {
			var a = e[t];
			return typeof a == "number" ? m(a) : a
		};
		var i = function(e, r, t) {
			var a = r.charCodeAt(0) - 32 << 4 | t.charCodeAt(0) - 48;
			return a == 59 ? e: m(a)
		};
		e["|"] = 254;
		var s = function(e) {
			return e.replace(/\n/g, " :").replace(/\r/g, " =")
		};
		function f(e, r) {
			switch (r.type) {
			case "base64":
				return o(_(e), r);
			case "binary":
				return o(e, r);
			case "buffer":
				return o(S && Buffer.isBuffer(e) ? e.toString("binary") : N(e), r);
			case "array":
				return o(kr(e), r);
			}
			throw new Error("Unrecognized type " + r.type)
		}
		function o(e, t) {
			var s = e.split(/[\n\r]+/),
			f = -1,
			o = -1,
			c = 0,
			u = 0,
			h = [];
			var d = [];
			var v = null;
			var p = {},
			m = [],
			b = [],
			g = [];
			var w = 0,
			k;
			var T = {
				Workbook: {
					WBProps: {},
					Names: []
				}
			};
			if ( + t.codepage >= 0) l( + t.codepage);
			for (; c !== s.length; ++c) {
				w = 0;
				var y = s[c].trim().replace(/\x1B([\x20-\x2F])([\x30-\x3F])/g, i).replace(r, n);
				var E = y.replace(/;;/g, "\0").split(";").map(function(e) {
					return e.replace(/\u0000/g, ";")
				});
				var _ = E[0],
				S;
				if (y.length > 0) switch (_) {
				case "ID":
					break;
				case "E":
					break;
				case "B":
					break;
				case "O":
					for (u = 1; u < E.length; ++u) switch (E[u].charAt(0)) {
					case "V":
						{
							var x = parseInt(E[u].slice(1), 10);
							if (x >= 1 && x <= 4) T.Workbook.WBProps.date1904 = true
						}
						break;
					}
					break;
				case "W":
					break;
				case "P":
					switch (E[1].charAt(0)) {
					case "P":
						d.push(y.slice(3).replace(/;;/g, ";"));
						break;
					}
					break;
				case "NN":
					{
						var A = {
							Sheet: 0
						};
						for (u = 1; u < E.length; ++u) switch (E[u].charAt(0)) {
						case "N":
							A.Name = E[u].slice(1);
							break;
						case "E":
							A.Ref = (t && t.sheet || "Sheet1") + "!" + rh(E[u].slice(1));
							break;
						}
						T.Workbook.Names.push(A)
					}
					break;
				case "C":
					var C = false,
					R = false,
					O = false,
					I = false,
					N = -1,
					F = -1,
					D = "",
					P = "z";
					var L = "";
					for (u = 1; u < E.length; ++u) switch (E[u].charAt(0)) {
					case "A":
						L = E[u].slice(1);
						break;
					case "X":
						o = parseInt(E[u].slice(1), 10) - 1;
						R = true;
						break;
					case "Y":
						f = parseInt(E[u].slice(1), 10) - 1;
						if (!R) o = 0;
						for (k = h.length; k <= f; ++k) h[k] = [];
						break;
					case "K":
						S = E[u].slice(1);
						if (S.charAt(0) === '"') {
							S = S.slice(1, S.length - 1);
							P = "s"
						} else if (S === "TRUE" || S === "FALSE") {
							S = S === "TRUE";
							P = "b"
						} else if (!isNaN(Er(S))) {
							S = Er(S);
							P = "n";
							if (v !== null && Le(v) && t.cellDates) {
								S = vr(T.Workbook.WBProps.date1904 ? S + 1462 : S);
								P = typeof S == "number" ? "n": "d"
							}
						}
						if (typeof a !== "undefined" && typeof S == "string" && (t || {}).type != "string" && (t || {}).codepage) S = a.utils.decode(t.codepage, S);
						C = true;
						break;
					case "E":
						I = true;
						D = rh(E[u].slice(1), {
							r: f,
							c: o
						});
						break;
					case "S":
						O = true;
						break;
					case "G":
						break;
					case "R":
						N = parseInt(E[u].slice(1), 10) - 1;
						break;
					case "C":
						F = parseInt(E[u].slice(1), 10) - 1;
						break;
					default:
						if (t && t.WTF) throw new Error("SYLK bad record " + y);
					}
					if (C) {
						if (!h[f][o]) h[f][o] = {
							t: P,
							v: S
						};
						else {
							h[f][o].t = P;
							h[f][o].v = S
						}
						if (v) h[f][o].z = v;
						if (t.cellText !== false && v) h[f][o].w = ze(h[f][o].z, h[f][o].v, {
							date1904: T.Workbook.WBProps.date1904
						});
						v = null
					}
					if (O) {
						if (I) throw new Error("SYLK shared formula cannot have own formula");
						var M = N > -1 && h[N][F];
						if (!M || !M[1]) throw new Error("SYLK shared formula cannot find base");
						D = ih(M[1], {
							r: f - N,
							c: o - F
						})
					}
					if (D) {
						if (!h[f][o]) h[f][o] = {
							t: "n",
							f: D
						};
						else h[f][o].f = D
					}
					if (L) {
						if (!h[f][o]) h[f][o] = {
							t: "z"
						};
						h[f][o].c = [{
							a: "SheetJSYLK",
							t: L
						}]
					}
					break;
				case "F":
					var U = 0;
					for (u = 1; u < E.length; ++u) switch (E[u].charAt(0)) {
					case "X":
						o = parseInt(E[u].slice(1), 10) - 1; ++U;
						break;
					case "Y":
						f = parseInt(E[u].slice(1), 10) - 1;
						for (k = h.length; k <= f; ++k) h[k] = [];
						break;
					case "M":
						w = parseInt(E[u].slice(1), 10) / 20;
						break;
					case "F":
						break;
					case "G":
						break;
					case "P":
						v = d[parseInt(E[u].slice(1), 10)];
						break;
					case "S":
						break;
					case "D":
						break;
					case "N":
						break;
					case "W":
						g = E[u].slice(1).split(" ");
						for (k = parseInt(g[0], 10); k <= parseInt(g[1], 10); ++k) {
							w = parseInt(g[2], 10);
							b[k - 1] = w === 0 ? {
								hidden: true
							}: {
								wch: w
							}
						}
						break;
					case "C":
						o = parseInt(E[u].slice(1), 10) - 1;
						if (!b[o]) b[o] = {};
						break;
					case "R":
						f = parseInt(E[u].slice(1), 10) - 1;
						if (!m[f]) m[f] = {};
						if (w > 0) {
							m[f].hpt = w;
							m[f].hpx = lc(w)
						} else if (w === 0) m[f].hidden = true;
						break;
					default:
						if (t && t.WTF) throw new Error("SYLK bad record " + y);
					}
					if (U < 1) v = null;
					break;
				default:
					if (t && t.WTF) throw new Error("SYLK bad record " + y);
				}
			}
			if (m.length > 0) p["!rows"] = m;
			if (b.length > 0) p["!cols"] = b;
			b.forEach(function(e) {
				nc(e)
			});
			if (t && t.sheetRows) h = h.slice(0, t.sheetRows);
			return [h, p, T]
		}
		function c(e, r) {
			var t = f(e, r);
			var a = t[0],
			n = t[1],
			i = t[2];
			var s = Tr(r);
			s.date1904 = (((i || {}).Workbook || {}).WBProps || {}).date1904;
			var l = qa(a, s);
			ir(n).forEach(function(e) {
				l[e] = n[e]
			});
			var o = Ya(l, r);
			ir(i).forEach(function(e) {
				o[e] = i[e]
			});
			o.bookType = "sylk";
			return o
		}
		function u(e, r, t, a, n, i) {
			var s = "C;Y" + (t + 1) + ";X" + (a + 1) + ";K";
			switch (e.t) {
			case "n":
				s += e.v || 0;
				if (e.f && !e.F) s += ";E" + nh(e.f, {
					r: t,
					c: a
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
				break;
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
					if (typeof r.width == "number" && !r.wpx) r.wpx = Qo(r.width);
					if (typeof r.wpx == "number" && !r.wch) r.wch = ec(r.wpx);
					if (typeof r.wch == "number") a += Math.round(r.wch)
				}
				if (a.charAt(a.length - 1) != " ") e.push(a)
			})
		}
		function v(e, r) {
			r.forEach(function(r, t) {
				var a = "F;";
				if (r.hidden) a += "M0;";
				else if (r.hpt) a += "M" + 20 * r.hpt + ";";
				else if (r.hpx) a += "M" + 20 * fc(r.hpx) + ";";
				if (a.length > 2) e.push(a + "R" + (t + 1))
			})
		}
		function p(e, r, t) {
			if (!r) r = {};
			r._formats = ["General"];
			var a = ["ID;PSheetJS;N;E"],
			n = [];
			var i = Ga(e["!ref"] || "A1"),
			s;
			var f = e["!data"] != null;
			var l = "\r\n";
			var o = (((t || {}).Workbook || {}).WBProps || {}).date1904;
			var c = "General";
			a.push("P;PGeneral");
			var p = i.s.r,
			m = i.s.c,
			b = [];
			if (e["!ref"]) for (p = i.s.r; p <= i.e.r; ++p) {
				if (f && !e["!data"][p]) continue;
				b = [];
				for (m = i.s.c; m <= i.e.c; ++m) {
					s = f ? e["!data"][p][m] : e[La(m) + Na(p)];
					if (!s || !s.c) continue;
					b.push(h(s.c, p, m))
				}
				if (b.length) n.push(b.join(l))
			}
			if (e["!ref"]) for (p = i.s.r; p <= i.e.r; ++p) {
				if (f && !e["!data"][p]) continue;
				b = [];
				for (m = i.s.c; m <= i.e.c; ++m) {
					s = f ? e["!data"][p][m] : e[La(m) + Na(p)];
					if (!s || s.v == null && (!s.f || s.F)) continue;
					if ((s.z || (s.t == "d" ? q[14] : "General")) != c) {
						var g = r._formats.indexOf(s.z);
						if (g == -1) {
							r._formats.push(s.z);
							g = r._formats.length - 1;
							a.push("P;P" + s.z.replace(/;/g, ";;"))
						}
						b.push("F;P" + g + ";Y" + (p + 1) + ";X" + (m + 1))
					}
					b.push(u(s, e, p, m, r, o))
				}
				n.push(b.join(l))
			}
			a.push("F;P0;DG0G8;M255");
			if (e["!cols"]) d(a, e["!cols"]);
			if (e["!rows"]) v(a, e["!rows"]);
			if (e["!ref"]) a.push("B;Y" + (i.e.r - i.s.r + 1) + ";X" + (i.e.c - i.s.c + 1) + ";D" + [i.s.c, i.s.r, i.e.c, i.e.r].join(" "));
			a.push("O;L;D;B" + (o ? ";V4": "") + ";K47;G100 0.001");
			delete r._formats;
			return a.join(l) + l + n.join(l) + l + "E" + l
		}
		return {
			to_workbook: c,
			from_sheet: p
		}
	} ();
	var jl = function() {
		function e(e, t) {
			switch (t.type) {
			case "base64":
				return r(_(e), t);
			case "binary":
				return r(e, t);
			case "buffer":
				return r(S && Buffer.isBuffer(e) ? e.toString("binary") : N(e), t);
			case "array":
				return r(kr(e), t);
			}
			throw new Error("Unrecognized type " + t.type)
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
				var f = t[i].trim().split(",");
				var l = f[0],
				o = f[1]; ++i;
				var c = t[i] || "";
				while ((c.match(/["]/g) || []).length & 1 && i < t.length - 1) c += "\n" + t[++i];
				c = c.trim();
				switch ( + l) {
				case - 1 : if (c === "BOT") {
						s[++a] = [];
						n = 0;
						continue
					} else if (c !== "EOD") throw new Error("Unrecognized DIF special command " + c);
					break;
				case 0:
					if (c === "TRUE") s[a][n] = true;
					else if (c === "FALSE") s[a][n] = false;
					else if (!isNaN(Er(o))) s[a][n] = Er(o);
					else if (!isNaN(Ir(o).getDate())) {
						s[a][n] = wr(o);
						if (! (r && r.UTC)) {
							s[a][n] = Fr(s[a][n])
						}
					} else s[a][n] = o; ++n;
					break;
				case 1:
					c = c.slice(1, c.length - 1);
					c = c.replace(/""/g, '"');
					if (w && c && c.match(/^=".*"$/)) c = c.slice(2, -1);
					s[a][n++] = c !== "" ? c: null;
					break;
				}
				if (c === "EOD") break
			}
			if (r && r.sheetRows) s = s.slice(0, r.sheetRows);
			return s
		}
		function t(r, t) {
			return qa(e(r, t), t)
		}
		function a(e, r) {
			var a = Ya(t(e, r), r);
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
			var t = Ga(e["!ref"]);
			var a = e["!data"] != null;
			var s = ['TABLE\r\n0,1\r\n"sheetjs"\r\n', "VECTORS\r\n0," + (t.e.r - t.s.r + 1) + '\r\n""\r\n', "TUPLES\r\n0," + (t.e.c - t.s.c + 1) + '\r\n""\r\n', 'DATA\r\n0,0\r\n""\r\n'];
			for (var f = t.s.r; f <= t.e.r; ++f) {
				var l = a ? e["!data"][f] : [];
				var o = "-1,0\r\nBOT\r\n";
				for (var c = t.s.c; c <= t.e.c; ++c) {
					var u = a ? l && l[c] : e[za({
						r: f,
						c: c
					})];
					if (u == null) {
						o += '1,0\r\n""\r\n';
						continue
					}
					switch (u.t) {
					case "n":
						if (r) {
							if (u.w != null) o += "0," + u.w + "\r\nV";
							else if (u.v != null) o += n(u.v, "V");
							else if (u.f != null && !u.F) o += i("=" + u.f);
							else o += '1,0\r\n""'
						} else {
							if (u.v == null) o += '1,0\r\n""';
							else o += n(u.v, "V")
						}
						break;
					case "b":
						o += u.v ? n(1, "TRUE") : n(0, "FALSE");
						break;
					case "s":
						o += i(!r || isNaN( + u.v) ? u.v: '="' + u.v + '"');
						break;
					case "d":
						if (!u.w) u.w = ze(u.z || q[14], dr(wr(u.v)));
						if (r) o += n(u.w, "V");
						else o += i(u.w);
						break;
					default:
						o += '1,0\r\n""';
					}
					o += "\r\n"
				}
				s.push(o)
			}
			return s.join("") + "-1,0\r\nEOD"
		}
		return {
			to_workbook: a,
			to_sheet: t,
			from_sheet: s
		}
	} ();
	var Kl = function() {
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
			f = [];
			for (; s !== a.length; ++s) {
				var l = a[s].trim().split(":");
				if (l[0] !== "cell") continue;
				var o = Wa(l[1]);
				if (f.length <= o.r) for (n = f.length; n <= o.r; ++n) if (!f[n]) f[n] = [];
				n = o.r;
				i = o.c;
				switch (l[2]) {
				case "t":
					f[n][i] = e(l[3]);
					break;
				case "v":
					f[n][i] = +l[3];
					break;
				case "vtf":
					var c = l[l.length - 1];
				case "vtc":
					switch (l[3]) {
					case "nl":
						f[n][i] = +l[4] ? true: false;
						break;
					default:
						f[n][i] = +l[4];
						break;
					}
					if (l[2] == "vtf") f[n][i] = [f[n][i], c];
				}
			}
			if (t && t.sheetRows) f = f.slice(0, t.sheetRows);
			return f
		}
		function a(e, r) {
			return qa(t(e, r), r)
		}
		function n(e, r) {
			return Ya(a(e, r), r)
		}
		var i = ["socialcalc:version:1.5", "MIME-Version: 1.0", "Content-Type: multipart/mixed; boundary=SocialCalcSpreadsheetControlSave"].join("\n");
		var s = ["--SocialCalcSpreadsheetControlSave", "Content-type: text/plain; charset=UTF-8"].join("\n") + "\n";
		var f = ["# SocialCalc Spreadsheet Control Save", "part:sheet"].join("\n");
		var l = "--SocialCalcSpreadsheetControlSave--";
		function o(e) {
			if (!e || !e["!ref"]) return "";
			var t = [],
			a = [],
			n,
			i = "";
			var s = Ha(e["!ref"]);
			var f = e["!data"] != null;
			for (var l = s.s.r; l <= s.e.r; ++l) {
				for (var o = s.s.c; o <= s.e.c; ++o) {
					i = za({
						r: l,
						c: o
					});
					n = f ? (e["!data"][l] || [])[o] : e[i];
					if (!n || n.v == null || n.t === "z") continue;
					a = ["cell", i, "t"];
					switch (n.t) {
					case "s":
						;
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
						var c = dr(wr(n.v));
						a[2] = "vtc";
						a[3] = "nd";
						a[4] = "" + c;
						a[5] = n.w || ze(n.z || q[14], c);
						break;
					case "e":
						continue;
					}
					t.push(a.join(":"))
				}
			}
			t.push("sheet:c:" + (s.e.c - s.s.c + 1) + ":r:" + (s.e.r - s.s.r + 1) + ":tvf:1");
			t.push("valueformat:1:text-wiki");
			return t.join("\n")
		}
		function c(e) {
			return [i, s, f, s, o(e), l].join("\n")
		}
		return {
			to_workbook: n,
			to_sheet: a,
			from_sheet: c
		}
	} ();
	var Yl = function() {
		function e(e, r, t, a, n) {
			if (n.raw) r[t][a] = e;
			else if (e === "") {} else if (e === "TRUE") r[t][a] = true;
			else if (e === "FALSE") r[t][a] = false;
			else if (!isNaN(Er(e))) r[t][a] = Er(e);
			else if (!isNaN(Ir(e).getDate())) r[t][a] = wr(e);
			else r[t][a] = e
		}
		function r(r, t) {
			var a = t || {};
			var n = [];
			if (!r || r.length === 0) return n;
			var i = r.split(/[\r\n]/);
			var s = i.length - 1;
			while (s >= 0 && i[s].length === 0)--s;
			var f = 10,
			l = 0;
			var o = 0;
			for (; o <= s; ++o) {
				l = i[o].indexOf(" ");
				if (l == -1) l = i[o].length;
				else l++;
				f = Math.max(f, l)
			}
			for (o = 0; o <= s; ++o) {
				n[o] = [];
				var c = 0;
				e(i[o].slice(0, f).trim(), n, o, c, a);
				for (c = 1; c <= (i[o].length - f) / 10 + 1; ++c) e(i[o].slice(f + (c - 1) * 10, f + c * 10).trim(), n, o, c, a)
			}
			if (a.sheetRows) n = n.slice(0, a.sheetRows);
			return n
		}
		var t = {
			44 : ",",
			9 : "\t",
			59 : ";",
			124 : "|"
		};
		var n = {
			44 : 3,
			9 : 2,
			59 : 1,
			124 : 0
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
			if (g != null && t.dense == null) t.dense = g;
			var n = {};
			if (t.dense) n["!data"] = [];
			var s = {
				s: {
					c: 0,
					r: 0
				},
				e: {
					c: 0,
					r: 0
				}
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
			var f = 0,
			l = 0,
			o = 0;
			var c = 0,
			u = 0,
			h = a.charCodeAt(0),
			d = false,
			v = 0,
			p = e.charCodeAt(0);
			var m = t.dateNF != null ? Ke(t.dateNF) : null;
			function b() {
				var r = e.slice(c, u);
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
					} else if (fh(r)) {
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
				} else if (!isNaN(o = Er(r))) {
					a.t = "n";
					a.v = o
				} else if (!isNaN((o = Ir(r)).getDate()) || m && r.match(m)) {
					a.z = t.dateNF || q[14];
					if (m && r.match(m)) {
						var i = Ye(r, t.dateNF, r.match(m) || []);
						o = wr(i);
						if (t && t.UTC === false) o = Fr(o)
					} else if (t && t.UTC === false) o = Fr(o);
					else if (t.cellText !== false && t.dateNF) a.w = ze(a.z, o);
					if (t.cellDates) {
						a.t = "d";
						a.v = o
					} else {
						a.t = "n";
						a.v = dr(o)
					}
					if (!t.cellNF) delete a.z
				} else {
					a.t = "s";
					a.v = r
				}
				if (a.t == "z") {} else if (t.dense) {
					if (!n["!data"][f]) n["!data"][f] = [];
					n["!data"][f][l] = a
				} else n[za({
					c: l,
					r: f
				})] = a;
				c = u + 1;
				p = e.charCodeAt(c);
				if (s.e.c < l) s.e.c = l;
				if (s.e.r < f) s.e.r = f;
				if (v == h)++l;
				else {
					l = 0; ++f;
					if (t.sheetRows && t.sheetRows <= f) return true
				}
			}
			e: for (; u < e.length; ++u) switch (v = e.charCodeAt(u)) {
			case 34:
				if (p === 34) d = !d;
				break;
			case 13:
				if (d) break;
				if (e.charCodeAt(u + 1) == 10)++u;
			case h:
				;
			case 10:
				if (!d && b()) break e;
				break;
			default:
				break;
			}
			if (u - c > 0) b();
			n["!ref"] = Va(s);
			return n
		}
		function f(e, t) {
			if (! (t && t.PRN)) return s(e, t);
			if (t.FS) return s(e, t);
			if (e.slice(0, 4) == "sep=") return s(e, t);
			if (e.indexOf("\t") >= 0 || e.indexOf(",") >= 0 || e.indexOf(";") >= 0) return s(e, t);
			return qa(r(e, t), t)
		}
		function l(e, r) {
			var t = "",
			n = r.type == "string" ? [0, 0, 0, 0] : uk(e, r);
			switch (r.type) {
			case "base64":
				t = _(e);
				break;
			case "binary":
				t = e;
				break;
			case "buffer":
				if (r.codepage == 65001) t = e.toString("utf8");
				else if (r.codepage && typeof a !== "undefined") t = a.utils.decode(r.codepage, e);
				else t = S && Buffer.isBuffer(e) ? e.toString("binary") : N(e);
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
			if (t.slice(0, 19) == "socialcalc:version:") return Kl.to_sheet(r.type == "string" ? t: kt(t), r);
			return f(t, r)
		}
		function o(e, r) {
			return Ya(l(e, r), r)
		}
		function c(e) {
			var r = [];
			if (!e["!ref"]) return "";
			var t = Ga(e["!ref"]),
			a;
			var n = e["!data"] != null;
			for (var i = t.s.r; i <= t.e.r; ++i) {
				var s = [];
				for (var f = t.s.c; f <= t.e.c; ++f) {
					var l = za({
						r: i,
						c: f
					});
					a = n ? (e["!data"][i] || [])[f] : e[l];
					if (!a || a.v == null) {
						s.push("          ");
						continue
					}
					var o = (a.w || (Ka(a), a.w) || "").slice(0, 10);
					while (o.length < 10) o += " ";
					s.push(o + (f === 0 ? " ": ""))
				}
				r.push(s.join(""))
			}
			return r.join("\n")
		}
		return {
			to_workbook: o,
			to_sheet: l,
			from_sheet: c
		}
	} ();
	function Zl(e, r) {
		var t = r || {},
		a = !!t.WTF;
		t.WTF = true;
		try {
			var n = Gl.to_workbook(e, t);
			t.WTF = a;
			return n
		} catch(i) {
			t.WTF = a;
			if (!i.message.match(/SYLK bad record ID/) && a) throw i;
			return Yl.to_workbook(e, r)
		}
	}
	var Jl = function() {
		function e(e, r, t) {
			if (!e) return;
			Ta(e, e.l || 0);
			var a = t.Enum || V;
			while (e.l < e.length) {
				var n = e._R(2);
				var i = a[n] || a[65535];
				var s = e._R(2);
				var f = e.l + s;
				var l = i.f && i.f(e, s, t);
				e.l = f;
				if (r(l, i, n)) return
			}
		}
		function r(e, r) {
			switch (r.type) {
			case "base64":
				return a(O(_(e)), r);
			case "binary":
				return a(O(e), r);
			case "buffer":
				;
			case "array":
				return a(e, r);
			}
			throw "Unsupported type " + r.type
		}
		var t = ["mmmm", "dd-mmm-yyyy", "dd-mmm", "mmm-yyyy", "@", "mm/dd", "hh:mm:ss AM/PM", "hh:mm AM/PM", "mm/dd/yyyy", "mm/dd", "hh:mm:ss", "hh:mm"];
		function a(r, a) {
			if (!r) return r;
			var n = a || {};
			if (g != null && n.dense == null) n.dense = g;
			var i = {},
			s = "Sheet1",
			f = "",
			l = 0;
			var o = {},
			c = [],
			u = [],
			h = [];
			if (n.dense) h = i["!data"] = [];
			var d = {
				s: {
					r: 0,
					c: 0
				},
				e: {
					r: 0,
					c: 0
				}
			};
			var v = n.sheetRows || 0;
			var p = {};
			if (r[4] == 81 && r[5] == 80 && r[6] == 87) return j(r, a);
			if (r[2] == 0) {
				if (r[3] == 8 || r[3] == 9) {
					if (r.length >= 16 && r[14] == 5 && r[15] === 108) throw new Error("Unsupported Works 3 for Mac file")
				}
			}
			if (r[2] == 2) {
				n.Enum = V;
				e(r,
				function(e, r, a) {
					switch (a) {
					case 0:
						n.vers = e;
						if (e >= 4096) n.qpro = true;
						break;
					case 255:
						n.vers = e;
						n.works = true;
						break;
					case 6:
						d = e;
						break;
					case 204:
						if (e) f = e;
						break;
					case 222:
						f = e;
						break;
					case 15:
						;
					case 51:
						if ((!n.qpro && !n.works || a == 51) && e[1].v.charCodeAt(0) < 48) e[1].v = e[1].v.slice(1);
						if (n.works || n.works2) e[1].v = e[1].v.replace(/\r\n/g, "\n");
					case 13:
						;
					case 14:
						;
					case 16:
						if ((e[2] & 112) == 112 && (e[2] & 15) > 1 && (e[2] & 15) < 15) {
							e[1].z = n.dateNF || t[(e[2] & 15) - 1] || q[14];
							if (n.cellDates) {
								e[1].v = vr(e[1].v);
								e[1].t = typeof e[1].v == "number" ? "n": "d"
							}
						}
						if (n.qpro) {
							if (e[3] > l) {
								i["!ref"] = Va(d);
								o[s] = i;
								c.push(s);
								i = {};
								if (n.dense) h = i["!data"] = [];
								d = {
									s: {
										r: 0,
										c: 0
									},
									e: {
										r: 0,
										c: 0
									}
								};
								l = e[3];
								s = f || "Sheet" + (l + 1);
								f = ""
							}
						}
						var u = n.dense ? (h[e[0].r] || [])[e[0].c] : i[za(e[0])];
						if (u) {
							u.t = e[1].t;
							u.v = e[1].v;
							if (e[1].z != null) u.z = e[1].z;
							if (e[1].f != null) u.f = e[1].f;
							p = u;
							break
						}
						if (n.dense) {
							if (!h[e[0].r]) h[e[0].r] = [];
							h[e[0].r][e[0].c] = e[1]
						} else i[za(e[0])] = e[1];
						p = e[1];
						break;
					case 21509:
						n.works2 = true;
						break;
					case 21506:
						{
							if (e == 5281) {
								p.z = "hh:mm:ss";
								if (n.cellDates && p.t == "n") {
									p.v = vr(p.v);
									p.t = typeof p.v == "number" ? "n": "d"
								}
							}
						}
						break;
					}
				},
				n)
			} else if (r[2] == 26 || r[2] == 14) {
				n.Enum = $;
				if (r[2] == 14) {
					n.qpro = true;
					r.l = 0
				}
				e(r,
				function(e, r, t) {
					switch (t) {
					case 204:
						s = e;
						break;
					case 22:
						if (e[1].v.charCodeAt(0) < 48) e[1].v = e[1].v.slice(1);
						e[1].v = e[1].v.replace(/\x0F./g,
						function(e) {
							return String.fromCharCode(e.charCodeAt(1) - 32)
						}).replace(/\r\n/g, "\n");
					case 23:
						;
					case 24:
						;
					case 25:
						;
					case 37:
						;
					case 39:
						;
					case 40:
						if (e[3] > l) {
							i["!ref"] = Va(d);
							o[s] = i;
							c.push(s);
							i = {};
							if (n.dense) h = i["!data"] = [];
							d = {
								s: {
									r: 0,
									c: 0
								},
								e: {
									r: 0,
									c: 0
								}
							};
							l = e[3];
							s = "Sheet" + (l + 1)
						}
						if (v > 0 && e[0].r >= v) break;
						if (n.dense) {
							if (!h[e[0].r]) h[e[0].r] = [];
							h[e[0].r][e[0].c] = e[1]
						} else i[za(e[0])] = e[1];
						if (d.e.c < e[0].c) d.e.c = e[0].c;
						if (d.e.r < e[0].r) d.e.r = e[0].r;
						break;
					case 27:
						if (e[14e3]) u[e[14e3][0]] = e[14e3][1];
						break;
					case 1537:
						u[e[0]] = e[1];
						if (e[0] == l) s = e[1];
						break;
					default:
						break;
					}
				},
				n)
			} else throw new Error("Unrecognized LOTUS BOF " + r[2]);
			i["!ref"] = Va(d);
			o[f || s] = i;
			c.push(f || s);
			if (!u.length) return {
				SheetNames: c,
				Sheets: o
			};
			var m = {},
			b = [];
			for (var w = 0; w < u.length; ++w) if (o[c[w]]) {
				b.push(u[w] || c[w]);
				m[u[w]] = o[u[w]] || o[c[w]]
			} else {
				b.push(u[w]);
				m[u[w]] = {
					"!ref": "A1"
				}
			}
			return {
				SheetNames: b,
				Sheets: m
			}
		}
		function n(e, r) {
			var t = r || {};
			if ( + t.codepage >= 0) l( + t.codepage);
			if (t.type == "string") throw new Error("Cannot write WK1 to JS string");
			var a = Sa();
			if (!e["!ref"]) throw new Error("Cannot export empty sheet to WK1");
			var n = Ga(e["!ref"]);
			var i = e["!data"] != null;
			var f = [];
			rg(a, 0, s(1030));
			rg(a, 6, c(n));
			var o = Math.min(n.e.r, 8191);
			for (var u = n.s.c; u <= n.e.c; ++u) f[u] = La(u);
			for (var h = n.s.r; h <= o; ++h) {
				var d = Na(h);
				for (u = n.s.c; u <= n.e.c; ++u) {
					var p = i ? (e["!data"][h] || [])[u] : e[f[u] + d];
					if (!p || p.t == "z") continue;
					switch (p.t) {
					case "n":
						if ((p.v | 0) == p.v && p.v >= -32768 && p.v <= 32767) rg(a, 13, b(h, u, p));
						else rg(a, 14, k(h, u, p));
						break;
					case "d":
						var m = dr(p.v);
						if ((m | 0) == m && m >= -32768 && m <= 32767) rg(a, 13, b(h, u, {
							t: "n",
							v: m,
							z: p.z || q[14]
						}));
						else rg(a, 14, k(h, u, {
							t: "n",
							v: m,
							z: p.z || q[14]
						}));
						break;
					default:
						var g = Ka(p);
						rg(a, 15, v(h, u, g.slice(0, 239)));
					}
				}
			}
			rg(a, 1);
			return a.end()
		}
		function i(e, r) {
			var t = r || {};
			if ( + t.codepage >= 0) l( + t.codepage);
			if (t.type == "string") throw new Error("Cannot write WK3 to JS string");
			var a = Sa();
			rg(a, 0, f(e));
			for (var n = 0,
			i = 0; n < e.SheetNames.length; ++n) if ((e.Sheets[e.SheetNames[n]] || {})["!ref"]) rg(a, 27, H(e.SheetNames[n], i++));
			var s = 0;
			for (n = 0; n < e.SheetNames.length; ++n) {
				var o = e.Sheets[e.SheetNames[n]];
				if (!o || !o["!ref"]) continue;
				var c = Ga(o["!ref"]);
				var u = o["!data"] != null;
				var h = [];
				var d = Math.min(c.e.r, 8191);
				for (var v = c.s.r; v <= d; ++v) {
					var p = Na(v);
					for (var m = c.s.c; m <= c.e.c; ++m) {
						if (v === c.s.r) h[m] = La(m);
						var b = h[m] + p;
						var g = u ? (o["!data"][v] || [])[m] : o[b];
						if (!g || g.t == "z") continue;
						if (g.t == "n") {
							rg(a, 23, F(v, m, s, g.v))
						} else {
							var w = Ka(g);
							rg(a, 22, R(v, m, s, w.slice(0, 239)))
						}
					}
				}++s
			}
			rg(a, 1);
			return a.end()
		}
		function s(e) {
			var r = Ea(2);
			r._W(2, e);
			return r
		}
		function f(e) {
			var r = Ea(26);
			r._W(2, 4096);
			r._W(2, 4);
			r._W(4, 0);
			var t = 0,
			a = 0,
			n = 0;
			for (var i = 0; i < e.SheetNames.length; ++i) {
				var s = e.SheetNames[i];
				var f = e.Sheets[s];
				if (!f || !f["!ref"]) continue; ++n;
				var l = Ha(f["!ref"]);
				if (t < l.e.r) t = l.e.r;
				if (a < l.e.c) a = l.e.c
			}
			if (t > 8191) t = 8191;
			r._W(2, t);
			r._W(1, n);
			r._W(1, a);
			r._W(2, 0);
			r._W(2, 0);
			r._W(1, 1);
			r._W(1, 2);
			r._W(4, 0);
			r._W(4, 0);
			return r
		}
		function o(e, r, t) {
			var a = {
				s: {
					c: 0,
					r: 0
				},
				e: {
					c: 0,
					r: 0
				}
			};
			if (r == 8 && t.qpro) {
				a.s.c = e._R(1);
				e.l++;
				a.s.r = e._R(2);
				a.e.c = e._R(1);
				e.l++;
				a.e.r = e._R(2);
				return a
			}
			a.s.c = e._R(2);
			a.s.r = e._R(2);
			if (r == 12 && t.qpro) e.l += 2;
			a.e.c = e._R(2);
			a.e.r = e._R(2);
			if (r == 12 && t.qpro) e.l += 2;
			if (a.s.c == 65535) a.s.c = a.e.c = a.s.r = a.e.r = 0;
			return a
		}
		function c(e) {
			var r = Ea(8);
			r._W(2, e.s.c);
			r._W(2, e.s.r);
			r._W(2, e.e.c);
			r._W(2, e.e.r);
			return r
		}
		function u(e, r, t) {
			var a = [{
				c: 0,
				r: 0
			},
			{
				t: "n",
				v: 0
			},
			0, 0];
			if (t.qpro && t.vers != 20768) {
				a[0].c = e._R(1);
				a[3] = e._R(1);
				a[0].r = e._R(2);
				e.l += 2
			} else if (t.works) {
				a[0].c = e._R(2);
				a[0].r = e._R(2);
				a[2] = e._R(2)
			} else {
				a[2] = e._R(1);
				a[0].c = e._R(2);
				a[0].r = e._R(2)
			}
			return a
		}
		function h(e) {
			if (e.z && Le(e.z)) {
				return 240 | (t.indexOf(e.z) + 1 || 2)
			}
			return 255
		}
		function d(e, r, t) {
			var a = e.l + r;
			var n = u(e, r, t);
			n[1].t = "s";
			if ((t.vers & 65534) == 20768) {
				e.l++;
				var i = e._R(1);
				n[1].v = e._R(i, "utf8");
				return n
			}
			if (t.qpro) e.l++;
			n[1].v = e._R(a - e.l, "cstr");
			return n
		}
		function v(e, r, t) {
			var a = Ea(7 + t.length);
			a._W(1, 255);
			a._W(2, r);
			a._W(2, e);
			a._W(1, 39);
			for (var n = 0; n < a.length; ++n) {
				var i = t.charCodeAt(n);
				a._W(1, i >= 128 ? 95 : i)
			}
			a._W(1, 0);
			return a
		}
		function p(e, r, t) {
			var a = e.l + r;
			var n = u(e, r, t);
			n[1].t = "s";
			if (t.vers == 20768) {
				var i = e._R(1);
				n[1].v = e._R(i, "utf8");
				return n
			}
			n[1].v = e._R(a - e.l, "cstr");
			return n
		}
		function m(e, r, t) {
			var a = u(e, r, t);
			a[1].v = e._R(2, "i");
			return a
		}
		function b(e, r, t) {
			var a = Ea(7);
			a._W(1, h(t));
			a._W(2, r);
			a._W(2, e);
			a._W(2, t.v, "i");
			return a
		}
		function w(e, r, t) {
			var a = u(e, r, t);
			a[1].v = e._R(8, "f");
			return a
		}
		function k(e, r, t) {
			var a = Ea(13);
			a._W(1, h(t));
			a._W(2, r);
			a._W(2, e);
			a._W(8, t.v, "f");
			return a
		}
		function T(e, r, t) {
			var a = e.l + r;
			var n = u(e, r, t);
			n[1].v = e._R(8, "f");
			if (t.qpro) e.l = a;
			else {
				var i = e._R(2);
				x(e.slice(e.l, e.l + i), n);
				e.l += i
			}
			return n
		}
		function y(e, r, t) {
			var a = r & 32768;
			r &= ~32768;
			r = (a ? e: 0) + (r >= 8192 ? r - 16384 : r);
			return (a ? "": "$") + (t ? La(r) : Na(r))
		}
		var E = {
			31 : ["NA", 0],
			33 : ["ABS", 1],
			34 : ["TRUNC", 1],
			35 : ["SQRT", 1],
			36 : ["LOG", 1],
			37 : ["LN", 1],
			38 : ["PI", 0],
			39 : ["SIN", 1],
			40 : ["COS", 1],
			41 : ["TAN", 1],
			42 : ["ATAN2", 2],
			43 : ["ATAN", 1],
			44 : ["ASIN", 1],
			45 : ["ACOS", 1],
			46 : ["EXP", 1],
			47 : ["MOD", 2],
			49 : ["ISNA", 1],
			50 : ["ISERR", 1],
			51 : ["FALSE", 0],
			52 : ["TRUE", 0],
			53 : ["RAND", 0],
			54 : ["DATE", 3],
			63 : ["ROUND", 2],
			64 : ["TIME", 3],
			68 : ["ISNUMBER", 1],
			69 : ["ISTEXT", 1],
			70 : ["LEN", 1],
			71 : ["VALUE", 1],
			73 : ["MID", 3],
			74 : ["CHAR", 1],
			80 : ["SUM", 69],
			81 : ["AVERAGEA", 69],
			82 : ["COUNTA", 69],
			83 : ["MINA", 69],
			84 : ["MAXA", 69],
			102 : ["UPPER", 1],
			103 : ["LOWER", 1],
			107 : ["PROPER", 1],
			109 : ["TRIM", 1],
			111 : ["T", 1]
		};
		var S = ["", "", "", "", "", "", "", "", "", "+", "-", "*", "/", "^", "=", "<>", "<=", ">=", "<", ">", "", "", "", "", "&", "", "", "", "", "", "", ""];
		function x(e, r) {
			Ta(e, 0);
			var t = [],
			a = 0,
			n = "",
			i = "",
			s = "",
			f = "";
			while (e.l < e.length) {
				var l = e[e.l++];
				switch (l) {
				case 0:
					t.push(e._R(8, "f"));
					break;
				case 1:
					{
						i = y(r[0].c, e._R(2), true);
						n = y(r[0].r, e._R(2), false);
						t.push(i + n)
					}
					break;
				case 2:
					{
						var o = y(r[0].c, e._R(2), true);
						var c = y(r[0].r, e._R(2), false);
						i = y(r[0].c, e._R(2), true);
						n = y(r[0].r, e._R(2), false);
						t.push(o + c + ":" + i + n)
					}
					break;
				case 3:
					if (e.l < e.length) {
						console.error("WK1 premature formula end");
						return
					}
					break;
				case 4:
					t.push("(" + t.pop() + ")");
					break;
				case 5:
					t.push(e._R(2));
					break;
				case 6:
					{
						var u = "";
						while (l = e[e.l++]) u += String.fromCharCode(l);
						t.push('"' + u.replace(/"/g, '""') + '"')
					}
					break;
				case 8:
					t.push("-" + t.pop());
					break;
				case 23:
					t.push("+" + t.pop());
					break;
				case 22:
					t.push("NOT(" + t.pop() + ")");
					break;
				case 20:
					;
				case 21:
					{
						f = t.pop();
						s = t.pop();
						t.push(["AND", "OR"][l - 20] + "(" + s + "," + f + ")")
					}
					break;
				default:
					if (l < 32 && S[l]) {
						f = t.pop();
						s = t.pop();
						t.push(s + S[l] + f)
					} else if (E[l]) {
						a = E[l][1];
						if (a == 69) a = e[e.l++];
						if (a > t.length) {
							console.error("WK1 bad formula parse 0x" + l.toString(16) + ":|" + t.join("|") + "|");
							return
						}
						var h = t.slice(-a);
						t.length -= a;
						t.push(E[l][0] + "(" + h.join(",") + ")")
					} else if (l <= 7) return console.error("WK1 invalid opcode " + l.toString(16));
					else if (l <= 24) return console.error("WK1 unsupported op " + l.toString(16));
					else if (l <= 30) return console.error("WK1 invalid opcode " + l.toString(16));
					else if (l <= 115) return console.error("WK1 unsupported function opcode " + l.toString(16));
					else return console.error("WK1 unrecognized opcode " + l.toString(16));
				}
			}
			if (t.length == 1) r[1].f = "" + t[0];
			else console.error("WK1 bad formula parse |" + t.join("|") + "|")
		}
		function A(e) {
			var r = [{
				c: 0,
				r: 0
			},
			{
				t: "n",
				v: 0
			},
			0];
			r[0].r = e._R(2);
			r[3] = e[e.l++];
			r[0].c = e[e.l++];
			return r
		}
		function C(e, r) {
			var t = A(e, r);
			t[1].t = "s";
			t[1].v = e._R(r - 4, "cstr");
			return t
		}
		function R(e, r, t, a) {
			var n = Ea(6 + a.length);
			n._W(2, e);
			n._W(1, t);
			n._W(1, r);
			n._W(1, 39);
			for (var i = 0; i < a.length; ++i) {
				var s = a.charCodeAt(i);
				n._W(1, s >= 128 ? 95 : s)
			}
			n._W(1, 0);
			return n
		}
		function I(e, r) {
			var t = A(e, r);
			t[1].v = e._R(2);
			var a = t[1].v >> 1;
			if (t[1].v & 1) {
				switch (a & 7) {
				case 0:
					a = (a >> 3) * 5e3;
					break;
				case 1:
					a = (a >> 3) * 500;
					break;
				case 2:
					a = (a >> 3) / 20;
					break;
				case 3:
					a = (a >> 3) / 200;
					break;
				case 4:
					a = (a >> 3) / 2e3;
					break;
				case 5:
					a = (a >> 3) / 2e4;
					break;
				case 6:
					a = (a >> 3) / 16;
					break;
				case 7:
					a = (a >> 3) / 64;
					break;
				}
			}
			t[1].v = a;
			return t
		}
		function N(e, r) {
			var t = A(e, r);
			var a = e._R(4);
			var n = e._R(4);
			var i = e._R(2);
			if (i == 65535) {
				if (a === 0 && n === 3221225472) {
					t[1].t = "e";
					t[1].v = 15
				} else if (a === 0 && n === 3489660928) {
					t[1].t = "e";
					t[1].v = 42
				} else t[1].v = 0;
				return t
			}
			var s = i & 32768;
			i = (i & 32767) - 16446;
			t[1].v = (1 - s * 2) * (n * Math.pow(2, i + 32) + a * Math.pow(2, i));
			return t
		}
		function F(e, r, t, a) {
			var n = Ea(14);
			n._W(2, e);
			n._W(1, t);
			n._W(1, r);
			if (a == 0) {
				n._W(4, 0);
				n._W(4, 0);
				n._W(2, 65535);
				return n
			}
			var i = 0,
			s = 0,
			f = 0,
			l = 0;
			if (a < 0) {
				i = 1;
				a = -a
			}
			s = Math.log2(a) | 0;
			a /= Math.pow(2, s - 31);
			l = a >>> 0;
			if ((l & 2147483648) == 0) {
				a /= 2; ++s;
				l = a >>> 0
			}
			a -= l;
			l |= 2147483648;
			l >>>= 0;
			a *= Math.pow(2, 32);
			f = a >>> 0;
			n._W(4, f);
			n._W(4, l);
			s += 16383 + (i ? 32768 : 0);
			n._W(2, s);
			return n
		}
		function D(e, r) {
			var t = N(e, 14);
			e.l += r - 14;
			return t
		}
		function P(e, r) {
			var t = A(e, r);
			var a = e._R(4);
			t[1].v = a >> 6;
			return t
		}
		function L(e, r) {
			var t = A(e, r);
			var a = e._R(8, "f");
			t[1].v = a;
			return t
		}
		function M(e, r) {
			var t = L(e, 12);
			e.l += r - 12;
			return t
		}
		function U(e, r) {
			return e[e.l + r - 1] == 0 ? e._R(r, "cstr") : ""
		}
		function B(e, r) {
			var t = e[e.l++];
			if (t > r - 1) t = r - 1;
			var a = "";
			while (a.length < t) a += String.fromCharCode(e[e.l++]);
			return a
		}
		function W(e, r, t) {
			if (!t.qpro || r < 21) return;
			var a = e._R(1);
			e.l += 17;
			e.l += 1;
			e.l += 2;
			var n = e._R(r - 21, "cstr");
			return [a, n]
		}
		function z(e, r) {
			var t = {},
			a = e.l + r;
			while (e.l < a) {
				var n = e._R(2);
				if (n == 14e3) {
					t[n] = [0, ""];
					t[n][0] = e._R(2);
					while (e[e.l]) {
						t[n][1] += String.fromCharCode(e[e.l]);
						e.l++
					}
					e.l++
				}
			}
			return t
		}
		function H(e, r) {
			var t = Ea(5 + e.length);
			t._W(2, 14e3);
			t._W(2, r);
			for (var a = 0; a < e.length; ++a) {
				var n = e.charCodeAt(a);
				t[t.l++] = n > 127 ? 95 : n
			}
			t[t.l++] = 0;
			return t
		}
		var V = {
			0 : {
				n: "BOF",
				f: ds
			},
			1 : {
				n: "EOF"
			},
			2 : {
				n: "CALCMODE"
			},
			3 : {
				n: "CALCORDER"
			},
			4 : {
				n: "SPLIT"
			},
			5 : {
				n: "SYNC"
			},
			6 : {
				n: "RANGE",
				f: o
			},
			7 : {
				n: "WINDOW1"
			},
			8 : {
				n: "COLW1"
			},
			9 : {
				n: "WINTWO"
			},
			10 : {
				n: "COLW2"
			},
			11 : {
				n: "NAME"
			},
			12 : {
				n: "BLANK"
			},
			13 : {
				n: "INTEGER",
				f: m
			},
			14 : {
				n: "NUMBER",
				f: w
			},
			15 : {
				n: "LABEL",
				f: d
			},
			16 : {
				n: "FORMULA",
				f: T
			},
			24 : {
				n: "TABLE"
			},
			25 : {
				n: "ORANGE"
			},
			26 : {
				n: "PRANGE"
			},
			27 : {
				n: "SRANGE"
			},
			28 : {
				n: "FRANGE"
			},
			29 : {
				n: "KRANGE1"
			},
			32 : {
				n: "HRANGE"
			},
			35 : {
				n: "KRANGE2"
			},
			36 : {
				n: "PROTEC"
			},
			37 : {
				n: "FOOTER"
			},
			38 : {
				n: "HEADER"
			},
			39 : {
				n: "SETUP"
			},
			40 : {
				n: "MARGINS"
			},
			41 : {
				n: "LABELFMT"
			},
			42 : {
				n: "TITLES"
			},
			43 : {
				n: "SHEETJS"
			},
			45 : {
				n: "GRAPH"
			},
			46 : {
				n: "NGRAPH"
			},
			47 : {
				n: "CALCCOUNT"
			},
			48 : {
				n: "UNFORMATTED"
			},
			49 : {
				n: "CURSORW12"
			},
			50 : {
				n: "WINDOW"
			},
			51 : {
				n: "STRING",
				f: p
			},
			55 : {
				n: "PASSWORD"
			},
			56 : {
				n: "LOCKED"
			},
			60 : {
				n: "QUERY"
			},
			61 : {
				n: "QUERYNAME"
			},
			62 : {
				n: "PRINT"
			},
			63 : {
				n: "PRINTNAME"
			},
			64 : {
				n: "GRAPH2"
			},
			65 : {
				n: "GRAPHNAME"
			},
			66 : {
				n: "ZOOM"
			},
			67 : {
				n: "SYMSPLIT"
			},
			68 : {
				n: "NSROWS"
			},
			69 : {
				n: "NSCOLS"
			},
			70 : {
				n: "RULER"
			},
			71 : {
				n: "NNAME"
			},
			72 : {
				n: "ACOMM"
			},
			73 : {
				n: "AMACRO"
			},
			74 : {
				n: "PARSE"
			},
			102 : {
				n: "PRANGES??"
			},
			103 : {
				n: "RRANGES??"
			},
			104 : {
				n: "FNAME??"
			},
			105 : {
				n: "MRANGES??"
			},
			204 : {
				n: "SHEETNAMECS",
				f: U
			},
			222 : {
				n: "SHEETNAMELP",
				f: B
			},
			255 : {
				n: "BOF",
				f: ds
			},
			21506 : {
				n: "WKSNF",
				f: ds
			},
			65535 : {
				n: ""
			}
		};
		var $ = {
			0 : {
				n: "BOF"
			},
			1 : {
				n: "EOF"
			},
			2 : {
				n: "PASSWORD"
			},
			3 : {
				n: "CALCSET"
			},
			4 : {
				n: "WINDOWSET"
			},
			5 : {
				n: "SHEETCELLPTR"
			},
			6 : {
				n: "SHEETLAYOUT"
			},
			7 : {
				n: "COLUMNWIDTH"
			},
			8 : {
				n: "HIDDENCOLUMN"
			},
			9 : {
				n: "USERRANGE"
			},
			10 : {
				n: "SYSTEMRANGE"
			},
			11 : {
				n: "ZEROFORCE"
			},
			12 : {
				n: "SORTKEYDIR"
			},
			13 : {
				n: "FILESEAL"
			},
			14 : {
				n: "DATAFILLNUMS"
			},
			15 : {
				n: "PRINTMAIN"
			},
			16 : {
				n: "PRINTSTRING"
			},
			17 : {
				n: "GRAPHMAIN"
			},
			18 : {
				n: "GRAPHSTRING"
			},
			19 : {
				n: "??"
			},
			20 : {
				n: "ERRCELL"
			},
			21 : {
				n: "NACELL"
			},
			22 : {
				n: "LABEL16",
				f: C
			},
			23 : {
				n: "NUMBER17",
				f: N
			},
			24 : {
				n: "NUMBER18",
				f: I
			},
			25 : {
				n: "FORMULA19",
				f: D
			},
			26 : {
				n: "FORMULA1A"
			},
			27 : {
				n: "XFORMAT",
				f: z
			},
			28 : {
				n: "DTLABELMISC"
			},
			29 : {
				n: "DTLABELCELL"
			},
			30 : {
				n: "GRAPHWINDOW"
			},
			31 : {
				n: "CPA"
			},
			32 : {
				n: "LPLAUTO"
			},
			33 : {
				n: "QUERY"
			},
			34 : {
				n: "HIDDENSHEET"
			},
			35 : {
				n: "??"
			},
			37 : {
				n: "NUMBER25",
				f: P
			},
			38 : {
				n: "??"
			},
			39 : {
				n: "NUMBER27",
				f: L
			},
			40 : {
				n: "FORMULA28",
				f: M
			},
			142 : {
				n: "??"
			},
			147 : {
				n: "??"
			},
			150 : {
				n: "??"
			},
			151 : {
				n: "??"
			},
			152 : {
				n: "??"
			},
			153 : {
				n: "??"
			},
			154 : {
				n: "??"
			},
			155 : {
				n: "??"
			},
			156 : {
				n: "??"
			},
			163 : {
				n: "??"
			},
			174 : {
				n: "??"
			},
			175 : {
				n: "??"
			},
			176 : {
				n: "??"
			},
			177 : {
				n: "??"
			},
			184 : {
				n: "??"
			},
			185 : {
				n: "??"
			},
			186 : {
				n: "??"
			},
			187 : {
				n: "??"
			},
			188 : {
				n: "??"
			},
			195 : {
				n: "??"
			},
			201 : {
				n: "??"
			},
			204 : {
				n: "SHEETNAMECS",
				f: U
			},
			205 : {
				n: "??"
			},
			206 : {
				n: "??"
			},
			207 : {
				n: "??"
			},
			208 : {
				n: "??"
			},
			256 : {
				n: "??"
			},
			259 : {
				n: "??"
			},
			260 : {
				n: "??"
			},
			261 : {
				n: "??"
			},
			262 : {
				n: "??"
			},
			263 : {
				n: "??"
			},
			265 : {
				n: "??"
			},
			266 : {
				n: "??"
			},
			267 : {
				n: "??"
			},
			268 : {
				n: "??"
			},
			270 : {
				n: "??"
			},
			271 : {
				n: "??"
			},
			384 : {
				n: "??"
			},
			389 : {
				n: "??"
			},
			390 : {
				n: "??"
			},
			393 : {
				n: "??"
			},
			396 : {
				n: "??"
			},
			512 : {
				n: "??"
			},
			514 : {
				n: "??"
			},
			513 : {
				n: "??"
			},
			516 : {
				n: "??"
			},
			517 : {
				n: "??"
			},
			640 : {
				n: "??"
			},
			641 : {
				n: "??"
			},
			642 : {
				n: "??"
			},
			643 : {
				n: "??"
			},
			644 : {
				n: "??"
			},
			645 : {
				n: "??"
			},
			646 : {
				n: "??"
			},
			647 : {
				n: "??"
			},
			648 : {
				n: "??"
			},
			658 : {
				n: "??"
			},
			659 : {
				n: "??"
			},
			660 : {
				n: "??"
			},
			661 : {
				n: "??"
			},
			662 : {
				n: "??"
			},
			665 : {
				n: "??"
			},
			666 : {
				n: "??"
			},
			768 : {
				n: "??"
			},
			772 : {
				n: "??"
			},
			1537 : {
				n: "SHEETINFOQP",
				f: W
			},
			1600 : {
				n: "??"
			},
			1602 : {
				n: "??"
			},
			1793 : {
				n: "??"
			},
			1794 : {
				n: "??"
			},
			1795 : {
				n: "??"
			},
			1796 : {
				n: "??"
			},
			1920 : {
				n: "??"
			},
			2048 : {
				n: "??"
			},
			2049 : {
				n: "??"
			},
			2052 : {
				n: "??"
			},
			2688 : {
				n: "??"
			},
			10998 : {
				n: "??"
			},
			12849 : {
				n: "??"
			},
			28233 : {
				n: "??"
			},
			28484 : {
				n: "??"
			},
			65535 : {
				n: ""
			}
		};
		var X = {
			5 : "dd-mmm-yy",
			6 : "dd-mmm",
			7 : "mmm-yy",
			8 : "mm/dd/yy",
			10 : "hh:mm:ss AM/PM",
			11 : "hh:mm AM/PM",
			14 : "dd-mmm-yyyy",
			15 : "mmm-yyyy",
			34 : "0.00",
			50 : "0.00;[Red]0.00",
			66 : "0.00;(0.00)",
			82 : "0.00;[Red](0.00)",
			162 : '"$"#,##0.00;\\("$"#,##0.00\\)',
			288 : "0%",
			304 : "0E+00",
			320 : "# ?/?"
		};
		function G(e) {
			var r = e._R(2);
			var t = e._R(1);
			if (t != 0) throw "unsupported QPW string type " + t.toString(16);
			return e._R(r, "sbcs-cont")
		}
		function j(e, r) {
			Ta(e, 0);
			var t = r || {};
			if (g != null && t.dense == null) t.dense = g;
			var a = {};
			if (t.dense) a["!data"] = [];
			var n = [],
			i = "",
			s = [];
			var f = {
				s: {
					r: -1,
					c: -1
				},
				e: {
					r: -1,
					c: -1
				}
			};
			var l = 0,
			o = 0,
			c = 0,
			u = 0;
			var h = {
				SheetNames: [],
				Sheets: {}
			};
			var d = [];
			e: while (e.l < e.length) {
				var v = e._R(2),
				p = e._R(2);
				var m = e.slice(e.l, e.l + p);
				Ta(m, 0);
				switch (v) {
				case 1:
					if (m._R(4) != 962023505) throw "Bad QPW9 BOF!";
					break;
				case 2:
					break e;
				case 8:
					break;
				case 10:
					{
						var b = m._R(4);
						var w = (m.length - m.l) / b | 0;
						for (var k = 0; k < b; ++k) {
							var T = m.l + w;
							var y = {};
							m.l += 2;
							y.numFmtId = m._R(2);
							if (X[y.numFmtId]) y.z = X[y.numFmtId];
							m.l = T;
							d.push(y)
						}
					}
					break;
				case 1025:
					break;
				case 1026:
					break;
				case 1031:
					{
						m.l += 12;
						while (m.l < m.length) {
							l = m._R(2);
							o = m._R(1);
							n.push(m._R(l, "cstr"))
						}
					}
					break;
				case 1032:
					{}
					break;
				case 1537:
					{
						var E = m._R(2);
						a = {};
						if (t.dense) a["!data"] = [];
						f.s.c = m._R(2);
						f.e.c = m._R(2);
						f.s.r = m._R(4);
						f.e.r = m._R(4);
						m.l += 4;
						if (m.l + 2 < m.length) {
							l = m._R(2);
							o = m._R(1);
							i = l == 0 ? "": m._R(l, "cstr")
						}
						if (!i) i = La(E)
					}
					break;
				case 1538:
					{
						if (f.s.c > 255 || f.s.r > 999999) break;
						if (f.e.c < f.s.c) f.e.c = f.s.c;
						if (f.e.r < f.s.r) f.e.r = f.s.r;
						a["!ref"] = Va(f);
						Kk(h, a, i)
					}
					break;
				case 2561:
					{
						c = m._R(2);
						if (f.e.c < c) f.e.c = c;
						if (f.s.c > c) f.s.c = c;
						u = m._R(4);
						if (f.s.r > u) f.s.r = u;
						u = m._R(4);
						if (f.e.r < u) f.e.r = u
					}
					break;
				case 3073:
					{
						u = m._R(4),
						l = m._R(4);
						if (f.s.r > u) f.s.r = u;
						if (f.e.r < u + l - 1) f.e.r = u + l - 1;
						var _ = La(c);
						while (m.l < m.length) {
							var S = {
								t: "z"
							};
							var x = m._R(1),
							A = -1;
							if (x & 128) A = m._R(2);
							var C = x & 64 ? m._R(2) - 1 : 0;
							switch (x & 31) {
							case 0:
								break;
							case 1:
								break;
							case 2:
								S = {
									t: "n",
									v: m._R(2)
								};
								break;
							case 3:
								S = {
									t: "n",
									v: m._R(2, "i")
								};
								break;
							case 4:
								S = {
									t: "n",
									v: Tn(m)
								};
								break;
							case 5:
								S = {
									t: "n",
									v: m._R(8, "f")
								};
								break;
							case 7:
								S = {
									t: "s",
									v: n[o = m._R(4) - 1]
								};
								break;
							case 8:
								S = {
									t: "n",
									v: m._R(8, "f")
								};
								m.l += 2;
								m.l += 4;
								if (isNaN(S.v)) S = {
									t: "e",
									v: 15
								};
								break;
							default:
								throw "Unrecognized QPW cell type " + (x & 31);
							}
							if (A != -1 && (d[A - 1] || {}).z) S.z = d[A - 1].z;
							var R = 0;
							if (x & 32) switch (x & 31) {
							case 2:
								R = m._R(2);
								break;
							case 3:
								R = m._R(2, "i");
								break;
							case 7:
								R = m._R(2);
								break;
							default:
								throw "Unsupported delta for QPW cell type " + (x & 31);
							}
							if (! (!t.sheetStubs && S.t == "z")) {
								var O = Tr(S);
								if (S.t == "n" && S.z && Le(S.z) && t.cellDates) {
									O.v = vr(S.v);
									O.t = typeof O.v == "number" ? "n": "d"
								}
								if (a["!data"] != null) {
									if (!a["!data"][u]) a["!data"][u] = [];
									a["!data"][u][c] = O
								} else a[_ + Na(u)] = O
							}++u; --l;
							while (C-->0 && l >= 0) {
								if (x & 32) switch (x & 31) {
								case 2:
									S = {
										t: "n",
										v: S.v + R & 65535
									};
									break;
								case 3:
									S = {
										t: "n",
										v: S.v + R & 65535
									};
									if (S.v > 32767) S.v -= 65536;
									break;
								case 7:
									S = {
										t: "s",
										v: n[o = o + R >>> 0]
									};
									break;
								default:
									throw "Cannot apply delta for QPW cell type " + (x & 31);
								} else switch (x & 31) {
								case 1:
									S = {
										t: "z"
									};
									break;
								case 2:
									S = {
										t: "n",
										v: m._R(2)
									};
									break;
								case 7:
									S = {
										t: "s",
										v: n[o = m._R(4) - 1]
									};
									break;
								default:
									throw "Cannot apply repeat for QPW cell type " + (x & 31);
								}
								if (A != -1);
								if (! (!t.sheetStubs && S.t == "z")) {
									if (a["!data"] != null) {
										if (!a["!data"][u]) a["!data"][u] = [];
										a["!data"][u][c] = S
									} else a[_ + Na(u)] = S
								}++u; --l
							}
						}
					}
					break;
				case 3074:
					{
						c = m._R(2);
						u = m._R(4);
						var I = G(m);
						if (a["!data"] != null) {
							if (!a["!data"][u]) a["!data"][u] = [];
							a["!data"][u][c] = {
								t: "s",
								v: I
							}
						} else a[La(c) + Na(u)] = {
							t: "s",
							v: I
						}
					}
					break;
				default:
					break;
				}
				e.l += p
			}
			return h
		}
		return {
			sheet_to_wk1: n,
			book_to_wk3: i,
			to_workbook: r
		}
	} ();
	function ql(e) {
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
				;
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
				;
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
				;
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
					break;
				};
			case "<u>":
				;
			case "<u/>":
				r.u = 1;
				break;
			case "</u>":
				break;
			case "<b":
				if (s.val == "0") break;
			case "<b>":
				;
			case "<b/>":
				r.b = 1;
				break;
			case "</b>":
				break;
			case "<i":
				if (s.val == "0") break;
			case "<i>":
				;
			case "<i/>":
				r.i = 1;
				break;
			case "</i>":
				break;
			case "<color":
				if (s.rgb) r.color = s.rgb.slice(2, 8);
				break;
			case "<color>":
				;
			case "<color/>":
				;
			case "</color>":
				break;
			case "<family":
				r.family = s.val;
				break;
			case "<family>":
				;
			case "<family/>":
				;
			case "</family>":
				break;
			case "<vertAlign":
				r.valign = s.val;
				break;
			case "<vertAlign>":
				;
			case "<vertAlign/>":
				;
			case "</vertAlign>":
				break;
			case "<scheme":
				break;
			case "<scheme>":
				;
			case "<scheme/>":
				;
			case "</scheme>":
				break;
			case "<extLst":
				;
			case "<extLst>":
				;
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
	var Ql = function() {
		var e = yt("t"),
		r = yt("rPr");
		function t(t) {
			var a = t.match(e);
			if (!a) return {
				t: "s",
				v: ""
			};
			var n = {
				t: "s",
				v: it(a[1])
			};
			var i = t.match(r);
			if (i) n.s = ql(i[1]);
			return n
		}
		var a = /<(?:\w+:)?r>/g,
		n = /<\/(?:\w+:)?r>/;
		return function i(e) {
			return e.replace(a, "").split(n).map(t).filter(function(e) {
				return e.v
			})
		}
	} ();
	var eo = function _T() {
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
	} ();
	var ro = /<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/g,
	to = /<(?:\w+:)?r\b[^>]*>/;
	var ao = /<(?:\w+:)?rPh.*?>([\s\S]*?)<\/(?:\w+:)?rPh>/g;
	function no(e, r) {
		var t = r ? r.cellHTML: true;
		var a = {};
		if (!e) return {
			t: ""
		};
		if (e.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
			a.t = it(kt(e.slice(e.indexOf(">") + 1).split(/<\/(?:\w+:)?t>/)[0] || ""), true);
			a.r = kt(e);
			if (t) a.h = ut(a.t)
		} else if (e.match(to)) {
			a.r = kt(e);
			a.t = it(kt((e.replace(ao, "").match(ro) || []).join("").replace(qr, "")), true);
			if (t) a.h = eo(Ql(a.r))
		}
		return a
	}
	var io = /<(?:\w+:)?sst([^>]*)>([\s\S]*)<\/(?:\w+:)?sst>/;
	var so = /<(?:\w+:)?(?:si|sstItem)>/g;
	var fo = /<\/(?:\w+:)?(?:si|sstItem)>/;
	function lo(e, r) {
		var t = [],
		a = "";
		if (!e) return t;
		var n = e.match(io);
		if (n) {
			a = n[2].replace(so, "").split(fo);
			for (var i = 0; i != a.length; ++i) {
				var s = no(a[i].trim(), r);
				if (s != null) t[t.length] = s
			}
			n = rt(n[1]);
			t.Count = n.count;
			t.Unique = n.uniqueCount
		}
		return t
	}
	var oo = /^\s|\s$|[\t\n\r]/;
	function co(e, r) {
		if (!r.bookSST) return "";
		var t = [Kr];
		t[t.length] = It("sst", null, {
			xmlns: Mt[0],
			count: e.Count,
			uniqueCount: e.Unique
		});
		for (var a = 0; a != e.length; ++a) {
			if (e[a] == null) continue;
			var n = e[a];
			var i = "<si>";
			if (n.r) i += n.r;
			else {
				i += "<t";
				if (!n.t) n.t = "";
				if (typeof n.t !== "string") n.t = String(n.t);
				if (n.t.match(oo)) i += ' xml:space="preserve"';
				i += ">" + lt(n.t) + "</t>"
			}
			i += "</si>";
			t[t.length] = i
		}
		if (t.length > 2) {
			t[t.length] = "</sst>";
			t[1] = t[1].replace("/>", ">")
		}
		return t.join("")
	}
	function uo(e) {
		return [e._R(4), e._R(4)]
	}
	function ho(e, r) {
		var t = [];
		var a = false;
		_a(e,
		function n(e, i, s) {
			switch (s) {
			case 159:
				t.Count = e[0];
				t.Unique = e[1];
				break;
			case 19:
				t.push(e);
				break;
			case 160:
				return true;
			case 35:
				a = true;
				break;
			case 36:
				a = false;
				break;
			default:
				if (i.T) {}
				if (!a || r.WTF) throw new Error("Unexpected record 0x" + s.toString(16));
			}
		});
		return t
	}
	function vo(e, r) {
		if (!r) r = Ea(8);
		r._W(4, e.Count);
		r._W(4, e.Unique);
		return r
	}
	var po = fn;
	function mo(e) {
		var r = Sa();
		xa(r, 159, vo(e));
		for (var t = 0; t < e.length; ++t) xa(r, 19, po(e[t]));
		xa(r, 160);
		return r.end()
	}
	function bo(e) {
		if (typeof a !== "undefined") return a.utils.encode(t, e);
		var r = [],
		n = e.split("");
		for (var i = 0; i < n.length; ++i) r[i] = n[i].charCodeAt(0);
		return r
	}
	function go(e, r) {
		var t = {};
		t.Major = e._R(2);
		t.Minor = e._R(2);
		if (r >= 4) e.l += r - 4;
		return t
	}
	function wo(e) {
		var r = {};
		r.id = e._R(0, "lpp4");
		r.R = go(e, 4);
		r.U = go(e, 4);
		r.W = go(e, 4);
		return r
	}
	function ko(e) {
		var r = e._R(4);
		var t = e.l + r - 4;
		var a = {};
		var n = e._R(4);
		var i = [];
		while (n-->0) i.push({
			t: e._R(4),
			v: e._R(0, "lpp4")
		});
		a.name = e._R(0, "lpp4");
		a.comps = i;
		if (e.l != t) throw new Error("Bad DataSpaceMapEntry: " + e.l + " != " + t);
		return a
	}
	function To(e) {
		var r = [];
		e.l += 4;
		var t = e._R(4);
		while (t-->0) r.push(ko(e));
		return r
	}
	function yo(e) {
		var r = [];
		e.l += 4;
		var t = e._R(4);
		while (t-->0) r.push(e._R(0, "lpp4"));
		return r
	}
	function Eo(e) {
		var r = {};
		e._R(4);
		e.l += 4;
		r.id = e._R(0, "lpp4");
		r.name = e._R(0, "lpp4");
		r.R = go(e, 4);
		r.U = go(e, 4);
		r.W = go(e, 4);
		return r
	}
	function _o(e) {
		var r = Eo(e);
		r.ename = e._R(0, "8lpp4");
		r.blksz = e._R(4);
		r.cmode = e._R(4);
		if (e._R(4) != 4) throw new Error("Bad !Primary record");
		return r
	}
	function So(e, r) {
		var t = e.l + r;
		var a = {};
		a.Flags = e._R(4) & 63;
		e.l += 4;
		a.AlgID = e._R(4);
		var n = false;
		switch (a.AlgID) {
		case 26126:
			;
		case 26127:
			;
		case 26128:
			n = a.Flags == 36;
			break;
		case 26625:
			n = a.Flags == 4;
			break;
		case 0:
			n = a.Flags == 16 || a.Flags == 4 || a.Flags == 36;
			break;
		default:
			throw "Unrecognized encryption algorithm: " + a.AlgID;
		}
		if (!n) throw new Error("Encryption Flags/AlgID mismatch");
		a.AlgIDHash = e._R(4);
		a.KeySize = e._R(4);
		a.ProviderType = e._R(4);
		e.l += 8;
		a.CSPName = e._R(t - e.l >> 1, "utf16le");
		e.l = t;
		return a
	}
	function xo(e, r) {
		var t = {},
		a = e.l + r;
		e.l += 4;
		t.Salt = e.slice(e.l, e.l + 16);
		e.l += 16;
		t.Verifier = e.slice(e.l, e.l + 16);
		e.l += 16;
		e._R(4);
		t.VerifierHash = e.slice(e.l, a);
		e.l = a;
		return t
	}
	function Ao(e) {
		var r = go(e);
		switch (r.Minor) {
		case 2:
			return [r.Minor, Co(e, r)];
		case 3:
			return [r.Minor, Ro(e, r)];
		case 4:
			return [r.Minor, Oo(e, r)];
		}
		throw new Error("ECMA-376 Encrypted file unrecognized Version: " + r.Minor)
	}
	function Co(e) {
		var r = e._R(4);
		if ((r & 63) != 36) throw new Error("EncryptionInfo mismatch");
		var t = e._R(4);
		var a = So(e, t);
		var n = xo(e, e.length - e.l);
		return {
			t: "Std",
			h: a,
			v: n
		}
	}
	function Ro() {
		throw new Error("File is password-protected: ECMA-376 Extensible")
	}
	function Oo(e) {
		var r = ["saltSize", "blockSize", "keyBits", "hashSize", "cipherAlgorithm", "cipherChaining", "hashAlgorithm", "saltValue"];
		e.l += 4;
		var t = e._R(e.length - e.l, "utf8");
		var a = {};
		t.replace(qr,
		function n(e) {
			var t = rt(e);
			switch (tt(t[0])) {
			case "<?xml":
				break;
			case "<encryption":
				;
			case "</encryption>":
				break;
			case "<keyData":
				r.forEach(function(e) {
					a[e] = t[e]
				});
				break;
			case "<dataIntegrity":
				a.encryptedHmacKey = t.encryptedHmacKey;
				a.encryptedHmacValue = t.encryptedHmacValue;
				break;
			case "<keyEncryptors>":
				;
			case "<keyEncryptors":
				a.encs = [];
				break;
			case "</keyEncryptors>":
				break;
			case "<keyEncryptor":
				a.uri = t.uri;
				break;
			case "</keyEncryptor>":
				break;
			case "<encryptedKey":
				a.encs.push(t);
				break;
			default:
				throw t[0];
			}
		});
		return a
	}
	function Io(e, r) {
		var t = {};
		var a = t.EncryptionVersionInfo = go(e, 4);
		r -= 4;
		if (a.Minor != 2) throw new Error("unrecognized minor version code: " + a.Minor);
		if (a.Major > 4 || a.Major < 2) throw new Error("unrecognized major version code: " + a.Major);
		t.Flags = e._R(4);
		r -= 4;
		var n = e._R(4);
		r -= 4;
		t.EncryptionHeader = So(e, n);
		r -= n;
		t.EncryptionVerifier = xo(e, r);
		return t
	}
	function No(e) {
		var r = {};
		var t = r.EncryptionVersionInfo = go(e, 4);
		if (t.Major != 1 || t.Minor != 1) throw "unrecognized version code " + t.Major + " : " + t.Minor;
		r.Salt = e._R(16);
		r.EncryptedVerifier = e._R(16);
		r.EncryptedVerifierHash = e._R(16);
		return r
	}
	function Fo(e) {
		var r = 0,
		t;
		var a = bo(e);
		var n = a.length + 1,
		i, s;
		var f, l, o;
		t = C(n);
		t[0] = a.length;
		for (i = 1; i != n; ++i) t[i] = a[i - 1];
		for (i = n - 1; i >= 0; --i) {
			s = t[i];
			f = (r & 16384) === 0 ? 0 : 1;
			l = r << 1 & 32767;
			o = f | l;
			r = o ^ s
		}
		return r ^ 52811
	}
	var Do = function() {
		var e = [187, 255, 255, 186, 255, 255, 185, 128, 0, 190, 15, 0, 191, 15, 0];
		var r = [57840, 7439, 52380, 33984, 4364, 3600, 61902, 12606, 6258, 57657, 54287, 34041, 10252, 43370, 20163];
		var t = [44796, 19929, 39858, 10053, 20106, 40212, 10761, 31585, 63170, 64933, 60267, 50935, 40399, 11199, 17763, 35526, 1453, 2906, 5812, 11624, 23248, 885, 1770, 3540, 7080, 14160, 28320, 56640, 55369, 41139, 20807, 41614, 21821, 43642, 17621, 28485, 56970, 44341, 19019, 38038, 14605, 29210, 60195, 50791, 40175, 10751, 21502, 43004, 24537, 18387, 36774, 3949, 7898, 15796, 31592, 63184, 47201, 24803, 49606, 37805, 14203, 28406, 56812, 17824, 35648, 1697, 3394, 6788, 13576, 27152, 43601, 17539, 35078, 557, 1114, 2228, 4456, 30388, 60776, 51953, 34243, 7079, 14158, 28316, 14128, 28256, 56512, 43425, 17251, 34502, 7597, 13105, 26210, 52420, 35241, 883, 1766, 3532, 4129, 8258, 16516, 33032, 4657, 9314, 18628];
		var a = function(e) {
			return (e / 2 | e * 128) & 255
		};
		var n = function(e, r) {
			return a(e ^ r)
		};
		var i = function(e) {
			var a = r[e.length - 1];
			var n = 104;
			for (var i = e.length - 1; i >= 0; --i) {
				var s = e[i];
				for (var f = 0; f != 7; ++f) {
					if (s & 64) a ^= t[n];
					s *= 2; --n
				}
			}
			return a
		};
		return function(r) {
			var t = bo(r);
			var a = i(t);
			var s = t.length;
			var f = C(16);
			for (var l = 0; l != 16; ++l) f[l] = 0;
			var o, c, u;
			if ((s & 1) === 1) {
				o = a >> 8;
				f[s] = n(e[0], o); --s;
				o = a & 255;
				c = t[t.length - 1];
				f[s] = n(c, o)
			}
			while (s > 0) {--s;
				o = a >> 8;
				f[s] = n(t[s], o); --s;
				o = a & 255;
				f[s] = n(t[s], o)
			}
			s = 15;
			u = 15 - t.length;
			while (u > 0) {
				o = a >> 8;
				f[s] = n(e[u], o); --s; --u;
				o = a & 255;
				f[s] = n(t[s], o); --s; --u
			}
			return f
		}
	} ();
	var Po = function(e, r, t, a, n) {
		if (!n) n = r;
		if (!a) a = Do(e);
		var i, s;
		for (i = 0; i != r.length; ++i) {
			s = r[i];
			s ^= a[t];
			s = (s >> 5 | s << 3) & 255;
			n[i] = s; ++t
		}
		return [n, t, a]
	};
	var Lo = function(e) {
		var r = 0,
		t = Do(e);
		return function(e) {
			var a = Po("", e, r, t);
			r = a[1];
			return a[0]
		}
	};
	function Mo(e, r, t, a) {
		var n = {
			key: ds(e),
			verificationBytes: ds(e)
		};
		if (t.password) n.verifier = Fo(t.password);
		a.valid = n.verificationBytes === n.verifier;
		if (a.valid) a.insitu = Lo(t.password);
		return n
	}
	function Uo(e, r, t) {
		var a = t || {};
		a.Info = e._R(2);
		e.l -= 2;
		if (a.Info === 1) a.Data = No(e, r);
		else a.Data = Io(e, r);
		return a
	}
	function Bo(e, r, t) {
		var a = {
			Type: t.biff >= 8 ? e._R(2) : 0
		};
		if (a.Type) Uo(e, r - 2, a);
		else Mo(e, t.biff >= 8 ? r: r - 2, t, a);
		return a
	}
	function Wo(e, r) {
		switch (r.type) {
		case "base64":
			return zo(_(e), r);
		case "binary":
			return zo(e, r);
		case "buffer":
			return zo(S && Buffer.isBuffer(e) ? e.toString("binary") : N(e), r);
		case "array":
			return zo(kr(e), r);
		}
		throw new Error("Unrecognized type " + r.type)
	}
	function zo(e, r) {
		var t = r || {};
		var a = {};
		var n = t.dense;
		if (n) a["!data"] = [];
		var i = e.match(/\\trowd[\s\S]*?\\row\b/g);
		if (!i) throw new Error("RTF missing table");
		var s = {
			s: {
				c: 0,
				r: 0
			},
			e: {
				c: 0,
				r: i.length - 1
			}
		};
		var f = [];
		i.forEach(function(e, r) {
			if (n) f = a["!data"][r] = [];
			var i = /\\[\w\-]+\b/g;
			var l = 0;
			var o;
			var c = -1;
			var u = [];
			while ((o = i.exec(e)) != null) {
				var h = e.slice(l, i.lastIndex - o[0].length);
				if (h.charCodeAt(0) == 32) h = h.slice(1);
				if (h.length) u.push(h);
				switch (o[0]) {
				case "\\cell":
					++c;
					if (u.length) {
						var d = {
							v: u.join(""),
							t: "s"
						};
						if (d.v == "TRUE" || d.v == "FALSE") {
							d.v = d.v == "TRUE";
							d.t = "b"
						} else if (!isNaN(Er(d.v))) {
							d.t = "n";
							if (t.cellText !== false) d.w = d.v;
							d.v = Er(d.v)
						}
						if (n) f[c] = d;
						else a[za({
							r: r,
							c: c
						})] = d
					}
					u = [];
					break;
				case "\\par":
					u.push("\n");
					break;
				}
				l = i.lastIndex
			}
			if (c > s.e.c) s.e.c = c
		});
		a["!ref"] = Va(s);
		return a
	}
	function Ho(e, r) {
		var t = Ya(Wo(e, r), r);
		t.bookType = "rtf";
		return t
	}
	function Vo(e, r) {
		var t = ["{\\rtf1\\ansi"];
		if (!e["!ref"]) return t[0] + "}";
		var a = Ga(e["!ref"]),
		n;
		var i = e["!data"] != null,
		s = [];
		for (var f = a.s.r; f <= a.e.r; ++f) {
			t.push("\\trowd\\trautofit1");
			for (var l = a.s.c; l <= a.e.c; ++l) t.push("\\cellx" + (l + 1));
			t.push("\\pard\\intbl");
			if (i) s = e["!data"][f] || [];
			for (l = a.s.c; l <= a.e.c; ++l) {
				var o = za({
					r: f,
					c: l
				});
				n = i ? s[l] : e[o];
				if (!n || n.v == null && (!n.f || n.F)) {
					t.push(" \\cell");
					continue
				}
				t.push(" " + (n.w || (Ka(n), n.w) || "").replace(/[\r\n]/g, "\\par "));
				t.push("\\cell")
			}
			t.push("\\pard\\intbl\\row")
		}
		return t.join("") + "}"
	}
	function $o(e) {
		var r = e.slice(e[0] === "#" ? 1 : 0).slice(0, 6);
		return [parseInt(r.slice(0, 2), 16), parseInt(r.slice(2, 4), 16), parseInt(r.slice(4, 6), 16)]
	}
	function Xo(e) {
		for (var r = 0,
		t = 1; r != 3; ++r) t = t * 256 + (e[r] > 255 ? 255 : e[r] < 0 ? 0 : e[r]);
		return t.toString(16).toUpperCase().slice(1)
	}
	function Go(e) {
		var r = e[0] / 255,
		t = e[1] / 255,
		a = e[2] / 255;
		var n = Math.max(r, t, a),
		i = Math.min(r, t, a),
		s = n - i;
		if (s === 0) return [0, 0, r];
		var f = 0,
		l = 0,
		o = n + i;
		l = s / (o > 1 ? 2 - o: o);
		switch (n) {
		case r:
			f = ((t - a) / s + 6) % 6;
			break;
		case t:
			f = (a - r) / s + 2;
			break;
		case a:
			f = (r - t) / s + 4;
			break;
		}
		return [f / 6, l, o / 2]
	}
	function jo(e) {
		var r = e[0],
		t = e[1],
		a = e[2];
		var n = t * 2 * (a < .5 ? a: 1 - a),
		i = a - n / 2;
		var s = [i, i, i],
		f = 6 * r;
		var l;
		if (t !== 0) switch (f | 0) {
		case 0:
			;
		case 6:
			l = n * f;
			s[0] += n;
			s[1] += l;
			break;
		case 1:
			l = n * (2 - f);
			s[0] += l;
			s[1] += n;
			break;
		case 2:
			l = n * (f - 2);
			s[1] += n;
			s[2] += l;
			break;
		case 3:
			l = n * (4 - f);
			s[1] += l;
			s[2] += n;
			break;
		case 4:
			l = n * (f - 4);
			s[2] += n;
			s[0] += l;
			break;
		case 5:
			l = n * (6 - f);
			s[2] += l;
			s[0] += n;
			break;
		}
		for (var o = 0; o != 3; ++o) s[o] = Math.round(s[o] * 255);
		return s
	}
	function Ko(e, r) {
		if (r === 0) return e;
		var t = Go($o(e));
		if (r < 0) t[2] = t[2] * (1 + r);
		else t[2] = 1 - (1 - t[2]) * (1 - r);
		return Xo(jo(t))
	}
	var Yo = 6,
	Zo = 15,
	Jo = 1,
	qo = Yo;
	function Qo(e) {
		return Math.floor((e + Math.round(128 / qo) / 256) * qo)
	}
	function ec(e) {
		return Math.floor((e - 5) / qo * 100 + .5) / 100
	}
	function rc(e) {
		return Math.round((e * qo + 5) / qo * 256) / 256
	}
	function tc(e) {
		return rc(ec(Qo(e)))
	}
	function ac(e) {
		var r = Math.abs(e - tc(e)),
		t = qo;
		if (r > .005) for (qo = Jo; qo < Zo; ++qo) if (Math.abs(e - tc(e)) <= r) {
			r = Math.abs(e - tc(e));
			t = qo
		}
		qo = t
	}
	function nc(e) {
		if (e.width) {
			e.wpx = Qo(e.width);
			e.wch = ec(e.wpx);
			e.MDW = qo
		} else if (e.wpx) {
			e.wch = ec(e.wpx);
			e.width = rc(e.wch);
			e.MDW = qo
		} else if (typeof e.wch == "number") {
			e.width = rc(e.wch);
			e.wpx = Qo(e.width);
			e.MDW = qo
		}
		if (e.customWidth) delete e.customWidth
	}
	var ic = 96,
	sc = ic;
	function fc(e) {
		return e * 96 / sc
	}
	function lc(e) {
		return e * sc / 96
	}
	var oc = {
		None: "none",
		Solid: "solid",
		Gray50: "mediumGray",
		Gray75: "darkGray",
		Gray25: "lightGray",
		HorzStripe: "darkHorizontal",
		VertStripe: "darkVertical",
		ReverseDiagStripe: "darkDown",
		DiagStripe: "darkUp",
		DiagCross: "darkGrid",
		ThickDiagCross: "darkTrellis",
		ThinHorzStripe: "lightHorizontal",
		ThinVertStripe: "lightVertical",
		ThinReverseDiagStripe: "lightDown",
		ThinHorzCross: "lightGrid"
	};
	function cc(e, r, t, a) {
		r.Borders = [];
		var n = {};
		var i = false; (e[0].match(qr) || []).forEach(function(e) {
			var t = rt(e);
			switch (tt(t[0])) {
			case "<borders":
				;
			case "<borders>":
				;
			case "</borders>":
				break;
			case "<border":
				;
			case "<border>":
				;
			case "<border/>":
				n = {};
				if (t.diagonalUp) n.diagonalUp = pt(t.diagonalUp);
				if (t.diagonalDown) n.diagonalDown = pt(t.diagonalDown);
				r.Borders.push(n);
				break;
			case "</border>":
				break;
			case "<left/>":
				break;
			case "<left":
				;
			case "<left>":
				break;
			case "</left>":
				break;
			case "<right/>":
				break;
			case "<right":
				;
			case "<right>":
				break;
			case "</right>":
				break;
			case "<top/>":
				break;
			case "<top":
				;
			case "<top>":
				break;
			case "</top>":
				break;
			case "<bottom/>":
				break;
			case "<bottom":
				;
			case "<bottom>":
				break;
			case "</bottom>":
				break;
			case "<diagonal":
				;
			case "<diagonal>":
				;
			case "<diagonal/>":
				break;
			case "</diagonal>":
				break;
			case "<horizontal":
				;
			case "<horizontal>":
				;
			case "<horizontal/>":
				break;
			case "</horizontal>":
				break;
			case "<vertical":
				;
			case "<vertical>":
				;
			case "<vertical/>":
				break;
			case "</vertical>":
				break;
			case "<start":
				;
			case "<start>":
				;
			case "<start/>":
				break;
			case "</start>":
				break;
			case "<end":
				;
			case "<end>":
				;
			case "<end/>":
				break;
			case "</end>":
				break;
			case "<color":
				;
			case "<color>":
				break;
			case "<color/>":
				;
			case "</color>":
				break;
			case "<extLst":
				;
			case "<extLst>":
				;
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
					if (!i) throw new Error("unrecognized " + t[0] + " in borders")
				};
			}
		})
	}
	function uc(e, r, t, a) {
		r.Fills = [];
		var n = {};
		var i = false; (e[0].match(qr) || []).forEach(function(e) {
			var t = rt(e);
			switch (tt(t[0])) {
			case "<fills":
				;
			case "<fills>":
				;
			case "</fills>":
				break;
			case "<fill>":
				;
			case "<fill":
				;
			case "<fill/>":
				n = {};
				r.Fills.push(n);
				break;
			case "</fill>":
				break;
			case "<gradientFill>":
				break;
			case "<gradientFill":
				;
			case "</gradientFill>":
				r.Fills.push(n);
				n = {};
				break;
			case "<patternFill":
				;
			case "<patternFill>":
				if (t.patternType) n.patternType = t.patternType;
				break;
			case "<patternFill/>":
				;
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
				;
			case "</bgColor>":
				break;
			case "<fgColor":
				if (!n.fgColor) n.fgColor = {};
				if (t.theme) n.fgColor.theme = parseInt(t.theme, 10);
				if (t.tint) n.fgColor.tint = parseFloat(t.tint);
				if (t.rgb != null) n.fgColor.rgb = t.rgb.slice(-6);
				break;
			case "<fgColor/>":
				;
			case "</fgColor>":
				break;
			case "<stop":
				;
			case "<stop/>":
				break;
			case "</stop>":
				break;
			case "<color":
				;
			case "<color/>":
				break;
			case "</color>":
				break;
			case "<extLst":
				;
			case "<extLst>":
				;
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
					if (!i) throw new Error("unrecognized " + t[0] + " in fills")
				};
			}
		})
	}
	function hc(e, r, t, a) {
		r.Fonts = [];
		var n = {};
		var s = false; (e[0].match(qr) || []).forEach(function(e) {
			var f = rt(e);
			switch (tt(f[0])) {
			case "<fonts":
				;
			case "<fonts>":
				;
			case "</fonts>":
				break;
			case "<font":
				;
			case "<font>":
				break;
			case "</font>":
				;
			case "<font/>":
				r.Fonts.push(n);
				n = {};
				break;
			case "<name":
				if (f.val) n.name = kt(f.val);
				break;
			case "<name/>":
				;
			case "</name>":
				break;
			case "<b":
				n.bold = f.val ? pt(f.val) : 1;
				break;
			case "<b/>":
				n.bold = 1;
				break;
			case "<i":
				n.italic = f.val ? pt(f.val) : 1;
				break;
			case "<i/>":
				n.italic = 1;
				break;
			case "<u":
				switch (f.val) {
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
					break;
				}
				break;
			case "<u/>":
				n.underline = 1;
				break;
			case "<strike":
				n.strike = f.val ? pt(f.val) : 1;
				break;
			case "<strike/>":
				n.strike = 1;
				break;
			case "<outline":
				n.outline = f.val ? pt(f.val) : 1;
				break;
			case "<outline/>":
				n.outline = 1;
				break;
			case "<shadow":
				n.shadow = f.val ? pt(f.val) : 1;
				break;
			case "<shadow/>":
				n.shadow = 1;
				break;
			case "<condense":
				n.condense = f.val ? pt(f.val) : 1;
				break;
			case "<condense/>":
				n.condense = 1;
				break;
			case "<extend":
				n.extend = f.val ? pt(f.val) : 1;
				break;
			case "<extend/>":
				n.extend = 1;
				break;
			case "<sz":
				if (f.val) n.sz = +f.val;
				break;
			case "<sz/>":
				;
			case "</sz>":
				break;
			case "<vertAlign":
				if (f.val) n.vertAlign = f.val;
				break;
			case "<vertAlign/>":
				;
			case "</vertAlign>":
				break;
			case "<family":
				if (f.val) n.family = parseInt(f.val, 10);
				break;
			case "<family/>":
				;
			case "</family>":
				break;
			case "<scheme":
				if (f.val) n.scheme = f.val;
				break;
			case "<scheme/>":
				;
			case "</scheme>":
				break;
			case "<charset":
				if (f.val == "1") break;
				f.codepage = i[parseInt(f.val, 10)];
				break;
			case "<color":
				if (!n.color) n.color = {};
				if (f.auto) n.color.auto = pt(f.auto);
				if (f.rgb) n.color.rgb = f.rgb.slice(-6);
				else if (f.indexed) {
					n.color.index = parseInt(f.indexed, 10);
					var l = ri[n.color.index];
					if (n.color.index == 81) l = ri[1];
					if (!l) l = ri[1];
					n.color.rgb = l[0].toString(16) + l[1].toString(16) + l[2].toString(16)
				} else if (f.theme) {
					n.color.theme = parseInt(f.theme, 10);
					if (f.tint) n.color.tint = parseFloat(f.tint);
					if (f.theme && t.themeElements && t.themeElements.clrScheme) {
						n.color.rgb = Ko(t.themeElements.clrScheme[n.color.theme].rgb, n.color.tint || 0)
					}
				}
				break;
			case "<color/>":
				;
			case "</color>":
				break;
			case "<AlternateContent":
				s = true;
				break;
			case "</AlternateContent>":
				s = false;
				break;
			case "<extLst":
				;
			case "<extLst>":
				;
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
					if (!s) throw new Error("unrecognized " + f[0] + " in fonts")
				};
			}
		})
	}
	function dc(e, r, t) {
		r.NumberFmt = [];
		var a = ir(q);
		for (var n = 0; n < a.length; ++n) r.NumberFmt[a[n]] = q[a[n]];
		var i = e[0].match(qr);
		if (!i) return;
		for (n = 0; n < i.length; ++n) {
			var s = rt(i[n]);
			switch (tt(s[0])) {
			case "<numFmts":
				;
			case "</numFmts>":
				;
			case "<numFmts/>":
				;
			case "<numFmts>":
				break;
			case "<numFmt":
				{
					var f = it(kt(s.formatCode)),
					l = parseInt(s.numFmtId, 10);
					r.NumberFmt[l] = f;
					if (l > 0) {
						if (l > 392) {
							for (l = 392; l > 60; --l) if (r.NumberFmt[l] == null) break;
							r.NumberFmt[l] = f
						}
						Je(f, l)
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
	function vc(e) {
		var r = ["<numFmts>"]; [[5, 8], [23, 26], [41, 44], [50, 392]].forEach(function(t) {
			for (var a = t[0]; a <= t[1]; ++a) if (e[a] != null) r[r.length] = It("numFmt", null, {
				numFmtId: a,
				formatCode: lt(e[a])
			})
		});
		if (r.length === 1) return "";
		r[r.length] = "</numFmts>";
		r[0] = It("numFmts", null, {
			count: r.length - 2
		}).replace("/>", ">");
		return r.join("")
	}
	var pc = ["numFmtId", "fillId", "fontId", "borderId", "xfId"];
	var mc = ["applyAlignment", "applyBorder", "applyFill", "applyFont", "applyNumberFormat", "applyProtection", "pivotButton", "quotePrefix"];
	function bc(e, r, t) {
		r.CellXf = [];
		var a;
		var n = false; (e[0].match(qr) || []).forEach(function(e) {
			var i = rt(e),
			s = 0;
			switch (tt(i[0])) {
			case "<cellXfs":
				;
			case "<cellXfs>":
				;
			case "<cellXfs/>":
				;
			case "</cellXfs>":
				break;
			case "<xf":
				;
			case "<xf/>":
				;
			case "<xf>":
				a = i;
				delete a[0];
				for (s = 0; s < pc.length; ++s) if (a[pc[s]]) a[pc[s]] = parseInt(a[pc[s]], 10);
				for (s = 0; s < mc.length; ++s) if (a[mc[s]]) a[mc[s]] = pt(a[mc[s]]);
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
				;
			case "<alignment/>":
				;
			case "<alignment>":
				var f = {};
				if (i.vertical) f.vertical = i.vertical;
				if (i.horizontal) f.horizontal = i.horizontal;
				if (i.textRotation != null) f.textRotation = i.textRotation;
				if (i.indent) f.indent = i.indent;
				if (i.wrapText) f.wrapText = pt(i.wrapText);
				a.alignment = f;
				break;
			case "</alignment>":
				break;
			case "<protection":
				;
			case "<protection>":
				break;
			case "</protection>":
				;
			case "<protection/>":
				break;
			case "<AlternateContent":
				;
			case "<AlternateContent>":
				n = true;
				break;
			case "</AlternateContent>":
				n = false;
				break;
			case "<extLst":
				;
			case "<extLst>":
				;
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
					if (!n) throw new Error("unrecognized " + i[0] + " in cellXfs")
				};
			}
		})
	}
	function gc(e) {
		var r = [];
		r[r.length] = It("cellXfs", null);
		e.forEach(function(e) {
			r[r.length] = It("xf", null, e)
		});
		r[r.length] = "</cellXfs>";
		if (r.length === 2) return "";
		r[0] = It("cellXfs", null, {
			count: r.length - 2
		}).replace("/>", ">");
		return r.join("")
	}
	var wc = function ST() {
		var e = /<(?:\w+:)?numFmts([^>]*)>[\S\s]*?<\/(?:\w+:)?numFmts>/;
		var r = /<(?:\w+:)?cellXfs([^>]*)>[\S\s]*?<\/(?:\w+:)?cellXfs>/;
		var t = /<(?:\w+:)?fills([^>]*)>[\S\s]*?<\/(?:\w+:)?fills>/;
		var a = /<(?:\w+:)?fonts([^>]*)>[\S\s]*?<\/(?:\w+:)?fonts>/;
		var n = /<(?:\w+:)?borders([^>]*)>[\S\s]*?<\/(?:\w+:)?borders>/;
		return function i(s, f, l) {
			var o = {};
			if (!s) return o;
			s = s.replace(/<!--([\s\S]*?)-->/gm, "").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm, "");
			var c;
			if (c = s.match(e)) dc(c, o, l);
			if (c = s.match(a)) hc(c, o, f, l);
			if (c = s.match(t)) uc(c, o, f, l);
			if (c = s.match(n)) cc(c, o, f, l);
			if (c = s.match(r)) bc(c, o, l);
			return o
		}
	} ();
	function kc(e, r) {
		var t = [Kr, It("styleSheet", null, {
			xmlns: Mt[0],
			"xmlns:vt": Lt.vt
		})],
		a;
		if (e.SSF && (a = vc(e.SSF)) != null) t[t.length] = a;
		t[t.length] = '<fonts count="1"><font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>';
		t[t.length] = '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>';
		t[t.length] = '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>';
		t[t.length] = '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
		if (a = gc(r.cellXfs)) t[t.length] = a;
		t[t.length] = '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
		t[t.length] = '<dxfs count="0"/>';
		t[t.length] = '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>';
		if (t.length > 2) {
			t[t.length] = "</styleSheet>";
			t[1] = t[1].replace("/>", ">")
		}
		return t.join("")
	}
	function Tc(e, r) {
		var t = e._R(2);
		var a = rn(e, r - 2);
		return [t, a]
	}
	function yc(e, r, t) {
		if (!t) t = Ea(6 + 4 * r.length);
		t._W(2, e);
		tn(r, t);
		var a = t.length > t.l ? t.slice(0, t.l) : t;
		if (t.l == null) t.l = t.length;
		return a
	}
	function Ec(e, r, t) {
		var a = {};
		a.sz = e._R(2) / 20;
		var n = In(e, 2, t);
		if (n.fItalic) a.italic = 1;
		if (n.fCondense) a.condense = 1;
		if (n.fExtend) a.extend = 1;
		if (n.fShadow) a.shadow = 1;
		if (n.fOutline) a.outline = 1;
		if (n.fStrikeout) a.strike = 1;
		var i = e._R(2);
		if (i === 700) a.bold = 1;
		switch (e._R(2)) {
		case 1:
			a.vertAlign = "superscript";
			break;
		case 2:
			a.vertAlign = "subscript";
			break;
		}
		var s = e._R(1);
		if (s != 0) a.underline = s;
		var f = e._R(1);
		if (f > 0) a.family = f;
		var l = e._R(1);
		if (l > 0) a.charset = l;
		e.l++;
		a.color = Rn(e, 8);
		switch (e._R(1)) {
		case 1:
			a.scheme = "major";
			break;
		case 2:
			a.scheme = "minor";
			break;
		}
		a.name = rn(e, r - 21);
		return a
	}
	function _c(e, r) {
		if (!r) r = Ea(25 + 4 * 32);
		r._W(2, e.sz * 20);
		Nn(e, r);
		r._W(2, e.bold ? 700 : 400);
		var t = 0;
		if (e.vertAlign == "superscript") t = 1;
		else if (e.vertAlign == "subscript") t = 2;
		r._W(2, t);
		r._W(1, e.underline || 0);
		r._W(1, e.family || 0);
		r._W(1, e.charset || 0);
		r._W(1, 0);
		On(e.color, r);
		var a = 0;
		if (e.scheme == "major") a = 1;
		if (e.scheme == "minor") a = 2;
		r._W(1, a);
		tn(e.name, r);
		return r.length > r.l ? r.slice(0, r.l) : r
	}
	var Sc = ["none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"];
	var xc;
	var Ac = ya;
	function Cc(e, r) {
		if (!r) r = Ea(4 * 3 + 8 * 7 + 16 * 1);
		if (!xc) xc = fr(Sc);
		var t = xc[e.patternType];
		if (t == null) t = 40;
		r._W(4, t);
		var a = 0;
		if (t != 40) {
			On({
				auto: 1
			},
			r);
			On({
				auto: 1
			},
			r);
			for (; a < 12; ++a) r._W(4, 0)
		} else {
			for (; a < 4; ++a) r._W(4, 0);
			for (; a < 12; ++a) r._W(4, 0)
		}
		return r.length > r.l ? r.slice(0, r.l) : r
	}
	function Rc(e, r) {
		var t = e.l + r;
		var a = e._R(2);
		var n = e._R(2);
		e.l = t;
		return {
			ixfe: a,
			numFmtId: n
		}
	}
	function Oc(e, r, t) {
		if (!t) t = Ea(16);
		t._W(2, r || 0);
		t._W(2, e.numFmtId || 0);
		t._W(2, 0);
		t._W(2, 0);
		t._W(2, 0);
		t._W(1, 0);
		t._W(1, 0);
		var a = 0;
		t._W(1, a);
		t._W(1, 0);
		t._W(1, 0);
		t._W(1, 0);
		return t
	}
	function Ic(e, r) {
		if (!r) r = Ea(10);
		r._W(1, 0);
		r._W(1, 0);
		r._W(4, 0);
		r._W(4, 0);
		return r
	}
	var Nc = ya;
	function Fc(e, r) {
		if (!r) r = Ea(51);
		r._W(1, 0);
		Ic(null, r);
		Ic(null, r);
		Ic(null, r);
		Ic(null, r);
		Ic(null, r);
		return r.length > r.l ? r.slice(0, r.l) : r
	}
	function Dc(e, r) {
		if (!r) r = Ea(12 + 4 * 10);
		r._W(4, e.xfId);
		r._W(2, 1);
		r._W(1, +e.builtinId);
		r._W(1, 0);
		bn(e.name || "", r);
		return r.length > r.l ? r.slice(0, r.l) : r
	}
	function Pc(e, r, t) {
		var a = Ea(4 + 256 * 2 * 4);
		a._W(4, e);
		bn(r, a);
		bn(t, a);
		return a.length > a.l ? a.slice(0, a.l) : a
	}
	function Lc(e, r, t) {
		var a = {};
		a.NumberFmt = [];
		for (var n in q) a.NumberFmt[n] = q[n];
		a.CellXf = [];
		a.Fonts = [];
		var i = [];
		var s = false;
		_a(e,
		function f(e, n, l) {
			switch (l) {
			case 44:
				a.NumberFmt[e[0]] = e[1];
				Je(e[1], e[0]);
				break;
			case 43:
				a.Fonts.push(e);
				if (e.color.theme != null && r && r.themeElements && r.themeElements.clrScheme) {
					e.color.rgb = Ko(r.themeElements.clrScheme[e.color.theme].rgb, e.color.tint || 0)
				}
				break;
			case 1025:
				break;
			case 45:
				break;
			case 46:
				break;
			case 47:
				if (i[i.length - 1] == 617) {
					a.CellXf.push(e)
				}
				break;
			case 48:
				;
			case 507:
				;
			case 572:
				;
			case 475:
				break;
			case 1171:
				;
			case 2102:
				;
			case 1130:
				;
			case 512:
				;
			case 2095:
				;
			case 3072:
				break;
			case 35:
				s = true;
				break;
			case 36:
				s = false;
				break;
			case 37:
				i.push(l);
				s = true;
				break;
			case 38:
				i.pop();
				s = false;
				break;
			default:
				if (n.T > 0) i.push(l);
				else if (n.T < 0) i.pop();
				else if (!s || t.WTF && i[i.length - 1] != 37) throw new Error("Unexpected record 0x" + l.toString(16));
			}
		});
		return a
	}
	function Mc(e, r) {
		if (!r) return;
		var t = 0; [[5, 8], [23, 26], [41, 44], [50, 392]].forEach(function(e) {
			for (var a = e[0]; a <= e[1]; ++a) if (r[a] != null)++t
		});
		if (t == 0) return;
		xa(e, 615, en(t)); [[5, 8], [23, 26], [41, 44], [50, 392]].forEach(function(t) {
			for (var a = t[0]; a <= t[1]; ++a) if (r[a] != null) xa(e, 44, yc(a, r[a]))
		});
		xa(e, 616)
	}
	function Uc(e) {
		var r = 1;
		if (r == 0) return;
		xa(e, 611, en(r));
		xa(e, 43, _c({
			sz: 12,
			color: {
				theme: 1
			},
			name: "Calibri",
			family: 2,
			scheme: "minor"
		}));
		xa(e, 612)
	}
	function Bc(e) {
		var r = 2;
		if (r == 0) return;
		xa(e, 603, en(r));
		xa(e, 45, Cc({
			patternType: "none"
		}));
		xa(e, 45, Cc({
			patternType: "gray125"
		}));
		xa(e, 604)
	}
	function Wc(e) {
		var r = 1;
		if (r == 0) return;
		xa(e, 613, en(r));
		xa(e, 46, Fc({}));
		xa(e, 614)
	}
	function zc(e) {
		var r = 1;
		xa(e, 626, en(r));
		xa(e, 47, Oc({
			numFmtId: 0,
			fontId: 0,
			fillId: 0,
			borderId: 0
		},
		65535));
		xa(e, 627)
	}
	function Hc(e, r) {
		xa(e, 617, en(r.length));
		r.forEach(function(r) {
			xa(e, 47, Oc(r, 0))
		});
		xa(e, 618)
	}
	function Vc(e) {
		var r = 1;
		xa(e, 619, en(r));
		xa(e, 48, Dc({
			xfId: 0,
			builtinId: 0,
			name: "Normal"
		}));
		xa(e, 620)
	}
	function $c(e) {
		var r = 0;
		xa(e, 505, en(r));
		xa(e, 506)
	}
	function Xc(e) {
		var r = 0;
		xa(e, 508, Pc(r, "TableStyleMedium9", "PivotStyleMedium4"));
		xa(e, 509)
	}
	function Gc() {
		return
	}
	function jc(e, r) {
		var t = Sa();
		xa(t, 278);
		Mc(t, e.SSF);
		Uc(t, e);
		Bc(t, e);
		Wc(t, e);
		zc(t, e);
		Hc(t, r.cellXfs);
		Vc(t, e);
		$c(t, e);
		Xc(t, e);
		Gc(t, e);
		xa(t, 279);
		return t.end()
	}
	var Kc = ["</a:lt1>", "</a:dk1>", "</a:lt2>", "</a:dk2>", "</a:accent1>", "</a:accent2>", "</a:accent3>", "</a:accent4>", "</a:accent5>", "</a:accent6>", "</a:hlink>", "</a:folHlink>"];
	function Yc(e, r, t) {
		r.themeElements.clrScheme = [];
		var a = {}; (e[0].match(qr) || []).forEach(function(e) {
			var n = rt(e);
			switch (n[0]) {
			case "<a:clrScheme":
				;
			case "</a:clrScheme>":
				break;
			case "<a:srgbClr":
				a.rgb = n.val;
				break;
			case "<a:sysClr":
				a.rgb = n.lastClr;
				break;
			case "<a:dk1>":
				;
			case "</a:dk1>":
				;
			case "<a:lt1>":
				;
			case "</a:lt1>":
				;
			case "<a:dk2>":
				;
			case "</a:dk2>":
				;
			case "<a:lt2>":
				;
			case "</a:lt2>":
				;
			case "<a:accent1>":
				;
			case "</a:accent1>":
				;
			case "<a:accent2>":
				;
			case "</a:accent2>":
				;
			case "<a:accent3>":
				;
			case "</a:accent3>":
				;
			case "<a:accent4>":
				;
			case "</a:accent4>":
				;
			case "<a:accent5>":
				;
			case "</a:accent5>":
				;
			case "<a:accent6>":
				;
			case "</a:accent6>":
				;
			case "<a:hlink>":
				;
			case "</a:hlink>":
				;
			case "<a:folHlink>":
				;
			case "</a:folHlink>":
				if (n[0].charAt(1) === "/") {
					r.themeElements.clrScheme[Kc.indexOf(n[0])] = a;
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
	function Zc() {}
	function Jc() {}
	var qc = /<a:clrScheme([^>]*)>[\s\S]*<\/a:clrScheme>/;
	var Qc = /<a:fontScheme([^>]*)>[\s\S]*<\/a:fontScheme>/;
	var eu = /<a:fmtScheme([^>]*)>[\s\S]*<\/a:fmtScheme>/;
	function ru(e, r, t) {
		r.themeElements = {};
		var a; [["clrScheme", qc, Yc], ["fontScheme", Qc, Zc], ["fmtScheme", eu, Jc]].forEach(function(n) {
			if (! (a = e.match(n[1]))) throw new Error(n[0] + " not found in themeElements");
			n[2](a, r, t)
		})
	}
	var tu = /<a:themeElements([^>]*)>[\s\S]*<\/a:themeElements>/;
	function au(e, r) {
		if (!e || e.length === 0) e = nu();
		var t;
		var a = {};
		if (! (t = e.match(tu))) throw new Error("themeElements not found in theme");
		ru(t[0], a, r);
		a.raw = e;
		return a
	}
	function nu(e, r) {
		if (r && r.themeXLSX) return r.themeXLSX;
		if (e && typeof e.raw == "string") return e.raw;
		var t = [Kr];
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
	function iu(e, r, t) {
		var a = e.l + r;
		var n = e._R(4);
		if (n === 124226) return;
		if (!t.cellStyles) {
			e.l = a;
			return
		}
		var i = e.slice(e.l);
		e.l = a;
		var s;
		try {
			s = Gr(i, {
				type: "array"
			})
		} catch(f) {
			return
		}
		var l = zr(s, "theme/theme/theme1.xml", true);
		if (!l) return;
		return au(l, t)
	}
	function su(e) {
		return e._R(4)
	}
	function fu(e) {
		var r = {};
		r.xclrType = e._R(2);
		r.nTintShade = e._R(2);
		switch (r.xclrType) {
		case 0:
			e.l += 4;
			break;
		case 1:
			r.xclrValue = lu(e, 4);
			break;
		case 2:
			r.xclrValue = Fs(e, 4);
			break;
		case 3:
			r.xclrValue = su(e, 4);
			break;
		case 4:
			e.l += 4;
			break;
		}
		e.l += 8;
		return r
	}
	function lu(e, r) {
		return ya(e, r)
	}
	function ou(e, r) {
		return ya(e, r)
	}
	function cu(e) {
		var r = e._R(2);
		var t = e._R(2) - 4;
		var a = [r];
		switch (r) {
		case 4:
			;
		case 5:
			;
		case 7:
			;
		case 8:
			;
		case 9:
			;
		case 10:
			;
		case 11:
			;
		case 13:
			a[1] = fu(e, t);
			break;
		case 6:
			a[1] = ou(e, t);
			break;
		case 14:
			;
		case 15:
			a[1] = e._R(t === 1 ? 1 : 2);
			break;
		default:
			throw new Error("Unrecognized ExtProp type: " + r + " " + t);
		}
		return a
	}
	function uu(e, r) {
		var t = e.l + r;
		e.l += 2;
		var a = e._R(2);
		e.l += 2;
		var n = e._R(2);
		var i = [];
		while (n-->0) i.push(cu(e, t - e.l));
		return {
			ixfe: a,
			ext: i
		}
	}
	function hu(e, r) {
		r.forEach(function(e) {
			switch (e[0]) {
			case 4:
				break;
			case 5:
				break;
			case 6:
				break;
			case 7:
				break;
			case 8:
				break;
			case 9:
				break;
			case 10:
				break;
			case 11:
				break;
			case 13:
				break;
			case 14:
				break;
			case 15:
				break;
			}
		})
	}
	function du(e, r) {
		return {
			flags: e._R(4),
			version: e._R(4),
			name: rn(e, r - 8)
		}
	}
	function vu(e) {
		var r = Ea(12 + 2 * e.name.length);
		r._W(4, e.flags);
		r._W(4, e.version);
		tn(e.name, r);
		return r.slice(0, r.l)
	}
	function pu(e) {
		var r = [];
		var t = e._R(4);
		while (t-->0) r.push([e._R(4), e._R(4)]);
		return r
	}
	function mu(e) {
		var r = Ea(4 + 8 * e.length);
		r._W(4, e.length);
		for (var t = 0; t < e.length; ++t) {
			r._W(4, e[t][0]);
			r._W(4, e[t][1])
		}
		return r
	}
	function bu(e, r) {
		var t = Ea(8 + 2 * r.length);
		t._W(4, e);
		tn(r, t);
		return t.slice(0, t.l)
	}
	function gu(e) {
		e.l += 4;
		return e._R(4) != 0
	}
	function wu(e, r) {
		var t = Ea(8);
		t._W(4, e);
		t._W(4, r ? 1 : 0);
		return t
	}
	function ku(e, r, t) {
		var a = {
			Types: [],
			Cell: [],
			Value: []
		};
		var n = t || {};
		var i = [];
		var s = false;
		var f = 2;
		_a(e,
		function(e, r, t) {
			switch (t) {
			case 335:
				a.Types.push({
					name:
					e.name
				});
				break;
			case 51:
				e.forEach(function(e) {
					if (f == 1) a.Cell.push({
						type: a.Types[e[0] - 1].name,
						index: e[1]
					});
					else if (f == 0) a.Value.push({
						type: a.Types[e[0] - 1].name,
						index: e[1]
					})
				});
				break;
			case 337:
				f = e ? 1 : 0;
				break;
			case 338:
				f = 2;
				break;
			case 35:
				i.push(t);
				s = true;
				break;
			case 36:
				i.pop();
				s = false;
				break;
			default:
				if (r.T) {} else if (!s || n.WTF && i[i.length - 1] != 35) throw new Error("Unexpected record 0x" + t.toString(16));
			}
		});
		return a
	}
	function Tu() {
		var e = Sa();
		xa(e, 332);
		xa(e, 334, en(1));
		xa(e, 335, vu({
			name: "XLDAPR",
			version: 12e4,
			flags: 3496657072
		}));
		xa(e, 336);
		xa(e, 339, bu(1, "XLDAPR"));
		xa(e, 52);
		xa(e, 35, en(514));
		xa(e, 4096, en(0));
		xa(e, 4097, vs(1));
		xa(e, 36);
		xa(e, 53);
		xa(e, 340);
		xa(e, 337, wu(1, true));
		xa(e, 51, mu([[1, 0]]));
		xa(e, 338);
		xa(e, 333);
		return e.end()
	}
	function yu(e, r, t) {
		var a = {
			Types: [],
			Cell: [],
			Value: []
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
				;
			case "</metadata>":
				break;
			case "<metadataTypes":
				;
			case "</metadataTypes>":
				break;
			case "<metadataType":
				a.Types.push({
					name:
					r.name
				});
				break;
			case "</metadataType>":
				break;
			case "<futureMetadata":
				for (var f = 0; f < a.Types.length; ++f) if (a.Types[f].name == r.name) s = a.Types[f];
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
					index: +r.v
				});
				else if (i == 0) a.Value.push({
					type: a.Types[r.t - 1].name,
					index: +r.v
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
				;
			case "<extLst>":
				;
			case "</extLst>":
				;
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
	function Eu() {
		var e = [Kr];
		e.push('<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">\n  <metadataTypes count="1">\n    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>\n  </metadataTypes>\n  <futureMetadata name="XLDAPR" count="1">\n    <bk>\n      <extLst>\n        <ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">\n          <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>\n        </ext>\n      </extLst>\n    </bk>\n  </futureMetadata>\n  <cellMetadata count="1">\n    <bk>\n      <rc t="1" v="0"/>\n    </bk>\n  </cellMetadata>\n</metadata>');
		return e.join("")
	}
	function _u(e) {
		var r = [];
		if (!e) return r;
		var t = 1; (e.match(qr) || []).forEach(function(e) {
			var a = rt(e);
			switch (a[0]) {
			case "<?xml":
				break;
			case "<calcChain":
				;
			case "<calcChain>":
				;
			case "</calcChain>":
				break;
			case "<c":
				delete a[0];
				if (a.i) t = a.i;
				else a.i = t;
				r.push(a);
				break;
			}
		});
		return r
	}
	function Su(e) {
		var r = {};
		r.i = e._R(4);
		var t = {};
		t.r = e._R(4);
		t.c = e._R(4);
		r.r = za(t);
		var a = e._R(1);
		if (a & 2) r.l = "1";
		if (a & 8) r.a = "1";
		return r
	}
	function xu(e, r, t) {
		var a = [];
		var n = false;
		_a(e,
		function i(e, r, s) {
			switch (s) {
			case 63:
				a.push(e);
				break;
			default:
				if (r.T) {} else if (!n || t.WTF) throw new Error("Unexpected record 0x" + s.toString(16));
			}
		});
		return a
	}
	function Au() {}
	function Cu(e, r, t, a) {
		if (!e) return e;
		var n = a || {};
		var i = false,
		s = false;
		_a(e,
		function f(e, r, t) {
			if (s) return;
			switch (t) {
			case 359:
				;
			case 363:
				;
			case 364:
				;
			case 366:
				;
			case 367:
				;
			case 368:
				;
			case 369:
				;
			case 370:
				;
			case 371:
				;
			case 472:
				;
			case 577:
				;
			case 578:
				;
			case 579:
				;
			case 580:
				;
			case 581:
				;
			case 582:
				;
			case 583:
				;
			case 584:
				;
			case 585:
				;
			case 586:
				;
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
	function Ru(e, r) {
		if (!e) return "??";
		var t = (e.match(/<c:chart [^>]*r:id="([^"]*)"/) || ["", ""])[1];
		return r["!id"][t].Target
	}
	var Ou = /<(?:\w+:)?shape(?:[^\w][^>]*)?>([\s\S]*?)<\/(?:\w+:)?shape>/g;
	function Iu(e, r, t) {
		var a = 0; (e.match(Ou) || []).forEach(function(e) {
			var n = "";
			var i = true;
			var s = -1;
			var f = -1,
			l = -1;
			e.replace(qr,
			function(r, t) {
				var a = rt(r);
				switch (tt(a[0])) {
				case "<ClientData":
					if (a.ObjectType) n = a.ObjectType;
					break;
				case "<Visible":
					;
				case "<Visible/>":
					i = false;
					break;
				case "<Row":
					;
				case "<Row>":
					s = t + r.length;
					break;
				case "</Row>":
					f = +e.slice(s, t).trim();
					break;
				case "<Column":
					;
				case "<Column>":
					s = t + r.length;
					break;
				case "</Column>":
					l = +e.slice(s, t).trim();
					break;
				}
				return ""
			});
			switch (n) {
			case "Note":
				var o = Xk(r, f >= 0 && l >= 0 ? za({
					r: f,
					c: l
				}) : t[a].ref);
				if (o.c) {
					o.c.hidden = i
				}++a;
				break;
			}
		})
	}
	function Nu(e, r, t) {
		var a = [21600, 21600];
		var n = ["m0,0l0", a[1], a[0], a[1], a[0], "0xe"].join(",");
		var i = [It("xml", null, {
			"xmlns:v": Ut.v,
			"xmlns:o": Ut.o,
			"xmlns:x": Ut.x,
			"xmlns:mv": Ut.mv
		}).replace(/\/>/, ">"), It("o:shapelayout", It("o:idmap", null, {
			"v:ext": "edit",
			data: e
		}), {
			"v:ext": "edit"
		})];
		var s = 65536 * e;
		var f = r || [];
		if (f.length > 0) i.push(It("v:shapetype", [It("v:stroke", null, {
			joinstyle: "miter"
		}), It("v:path", null, {
			gradientshapeok: "t",
			"o:connecttype": "rect"
		})].join(""), {
			id: "_x0000_t202",
			coordsize: a.join(","),
			"o:spt": 202,
			path: n
		}));
		f.forEach(function(e) {++s;
			i.push(Fu(e, s))
		});
		i.push("</xml>");
		return i.join("")
	}
	function Fu(e, r, t) {
		var a = Wa(e[0]);
		var n = {
			color2: "#BEFF82",
			type: "gradient"
		};
		if (n.type == "gradient") n.angle = "-180";
		var i = n.type == "gradient" ? It("o:fill", null, {
			type: "gradientUnscaled",
			"v:ext": "view"
		}) : null;
		var s = It("v:fill", i, n);
		var f = {
			on: "t",
			obscured: "t"
		};
		return ["<v:shape" + Ot({
			id: "_x0000_s" + r,
			type: "#_x0000_t202",
			style: "position:absolute; margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:10" + (e[1].hidden ? ";visibility:hidden": ""),
			fillcolor: "#ECFAD4",
			strokecolor: "#edeaa1"
		}) + ">", s, It("v:shadow", null, f), It("v:path", null, {
			"o:connecttype": "none"
		}), '<v:textbox><div style="text-align:left"></div></v:textbox>', '<x:ClientData ObjectType="Note">', "<x:MoveWithCells/>", "<x:SizeWithCells/>", Rt("x:Anchor", [a.c + 1, 0, a.r + 1, 0, a.c + 3, 20, a.r + 5, 20].join(",")), Rt("x:AutoFill", "False"), Rt("x:Row", String(a.r)), Rt("x:Column", String(a.c)), e[1].hidden ? "": "<x:Visible/>", "</x:ClientData>", "</v:shape>"].join("")
	}
	function Du(e, r, t, a) {
		var n = e["!data"] != null;
		var i;
		r.forEach(function(r) {
			var s = Wa(r.ref);
			if (s.r < 0 || s.c < 0) return;
			if (n) {
				if (!e["!data"][s.r]) e["!data"][s.r] = [];
				i = e["!data"][s.r][s.c]
			} else i = e[r.ref];
			if (!i) {
				i = {
					t: "z"
				};
				if (n) e["!data"][s.r][s.c] = i;
				else e[r.ref] = i;
				var f = Ga(e["!ref"] || "BDWGO1000001:A1");
				if (f.s.r > s.r) f.s.r = s.r;
				if (f.e.r < s.r) f.e.r = s.r;
				if (f.s.c > s.c) f.s.c = s.c;
				if (f.e.c < s.c) f.e.c = s.c;
				var l = Va(f);
				e["!ref"] = l
			}
			if (!i.c) i.c = [];
			var o = {
				a: r.author,
				t: r.t,
				r: r.r,
				T: t
			};
			if (r.h) o.h = r.h;
			for (var c = i.c.length - 1; c >= 0; --c) {
				if (!t && i.c[c].T) return;
				if (t && !i.c[c].T) i.c.splice(c, 1)
			}
			if (t && a) for (c = 0; c < a.length; ++c) {
				if (o.a == a[c].id) {
					o.a = a[c].name || o.a;
					break
				}
			}
			i.c.push(o)
		})
	}
	function Pu(e, r) {
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
				author: i.authorId && t[i.authorId] || "sheetjsghost",
				ref: i.ref,
				guid: i.guid
			};
			var f = Wa(i.ref);
			if (r.sheetRows && r.sheetRows <= f.r) return;
			var l = e.match(/<(?:\w+:)?text>([\s\S]*)<\/(?:\w+:)?text>/);
			var o = !!l && !!l[1] && no(l[1]) || {
				r: "",
				t: "",
				h: ""
			};
			s.r = o.r;
			if (o.r == "<t></t>") o.t = o.h = "";
			s.t = (o.t || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
			if (r.cellHTML) s.h = o.h;
			a.push(s)
		});
		return a
	}
	function Lu(e) {
		var r = [Kr, It("comments", null, {
			xmlns: Mt[0]
		})];
		var t = [];
		r.push("<authors>");
		e.forEach(function(e) {
			e[1].forEach(function(e) {
				var a = lt(e.a);
				if (t.indexOf(a) == -1) {
					t.push(a);
					r.push("<author>" + a + "</author>")
				}
				if (e.T && e.ID && t.indexOf("tc=" + e.ID) == -1) {
					t.push("tc=" + e.ID);
					r.push("<author>" + "tc=" + e.ID + "</author>")
				}
			})
		});
		if (t.length == 0) {
			t.push("SheetJ5");
			r.push("<author>SheetJ5</author>")
		}
		r.push("</authors>");
		r.push("<commentList>");
		e.forEach(function(e) {
			var a = 0,
			n = [],
			i = 0;
			if (e[1][0] && e[1][0].T && e[1][0].ID) a = t.indexOf("tc=" + e[1][0].ID);
			e[1].forEach(function(e) {
				if (e.a) a = t.indexOf(lt(e.a));
				if (e.T)++i;
				n.push(e.t == null ? "": lt(e.t))
			});
			if (i === 0) {
				e[1].forEach(function(a) {
					r.push('<comment ref="' + e[0] + '" authorId="' + t.indexOf(lt(a.a)) + '"><text>');
					r.push(Rt("t", a.t == null ? "": lt(a.t)));
					r.push("</text></comment>")
				})
			} else {
				if (e[1][0] && e[1][0].T && e[1][0].ID) a = t.indexOf("tc=" + e[1][0].ID);
				r.push('<comment ref="' + e[0] + '" authorId="' + a + '"><text>');
				var s = "Comment:\n    " + n[0] + "\n";
				for (var f = 1; f < n.length; ++f) s += "Reply:\n    " + n[f] + "\n";
				r.push(Rt("t", lt(s)));
				r.push("</text></comment>")
			}
		});
		r.push("</commentList>");
		if (r.length > 2) {
			r[r.length] = "</comments>";
			r[1] = r[1].replace("/>", ">")
		}
		return r.join("")
	}
	function Mu(e, r) {
		var t = [];
		var a = false,
		n = {},
		i = 0;
		e.replace(qr,
		function s(f, l) {
			var o = rt(f);
			switch (tt(o[0])) {
			case "<?xml":
				break;
			case "<ThreadedComments":
				break;
			case "</ThreadedComments>":
				break;
			case "<threadedComment":
				n = {
					author: o.personId,
					guid: o.id,
					ref: o.ref,
					T: 1
				};
				break;
			case "</threadedComment>":
				if (n.t != null) t.push(n);
				break;
			case "<text>":
				;
			case "<text":
				i = l + f.length;
				break;
			case "</text>":
				n.t = e.slice(i, l).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
				break;
			case "<mentions":
				;
			case "<mentions>":
				a = true;
				break;
			case "</mentions>":
				a = false;
				break;
			case "<extLst":
				;
			case "<extLst>":
				;
			case "</extLst>":
				;
			case "<extLst/>":
				break;
			case "<ext":
				a = true;
				break;
			case "</ext>":
				a = false;
				break;
			default:
				if (!a && r.WTF) throw new Error("unrecognized " + o[0] + " in threaded comments");
			}
			return f
		});
		return t
	}
	function Uu(e, r, t) {
		var a = [Kr, It("ThreadedComments", null, {
			xmlns: Lt.TCMNT
		}).replace(/[\/]>/, ">")];
		e.forEach(function(e) {
			var n = ""; (e[1] || []).forEach(function(i, s) {
				if (!i.T) {
					delete i.ID;
					return
				}
				if (i.a && r.indexOf(i.a) == -1) r.push(i.a);
				var f = {
					ref: e[0],
					id: "{54EE7951-7262-4200-6969-" + ("000000000000" + t.tcid++).slice(-12) + "}"
				};
				if (s == 0) n = f.id;
				else f.parentId = n;
				i.ID = f.id;
				if (i.a) f.personId = "{54EE7950-7262-4200-6969-" + ("000000000000" + r.indexOf(i.a)).slice(-12) + "}";
				a.push(It("threadedComment", Rt("text", i.t || ""), f))
			})
		});
		a.push("</ThreadedComments>");
		return a.join("")
	}
	function Bu(e, r) {
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
					id: n.id
				});
				break;
			case "</person>":
				break;
			case "<extLst":
				;
			case "<extLst>":
				;
			case "</extLst>":
				;
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
	function Wu(e) {
		var r = [Kr, It("personList", null, {
			xmlns: Lt.TCMNT,
			"xmlns:x": Mt[0]
		}).replace(/[\/]>/, ">")];
		e.forEach(function(e, t) {
			r.push(It("person", null, {
				displayName: e,
				id: "{54EE7950-7262-4200-6969-" + ("000000000000" + t).slice(-12) + "}",
				userId: e,
				providerId: "None"
			}))
		});
		r.push("</personList>");
		return r.join("")
	}
	function zu(e) {
		var r = {};
		r.iauthor = e._R(4);
		var t = Sn(e, 16);
		r.rfx = t.s;
		r.ref = za(t.s);
		e.l += 16;
		return r
	}
	function Hu(e, r) {
		if (r == null) r = Ea(36);
		r._W(4, e[1].iauthor);
		xn(e[0], r);
		r._W(4, 0);
		r._W(4, 0);
		r._W(4, 0);
		r._W(4, 0);
		return r
	}
	var Vu = rn;
	function $u(e) {
		return tn(e.slice(0, 54))
	}
	function Xu(e, r) {
		var t = [];
		var a = [];
		var n = {};
		var i = false;
		_a(e,
		function s(e, f, l) {
			switch (l) {
			case 632:
				a.push(e);
				break;
			case 635:
				n = e;
				break;
			case 637:
				n.t = e.t;
				n.h = e.h;
				n.r = e.r;
				break;
			case 636:
				n.author = a[n.iauthor];
				delete n.iauthor;
				if (r.sheetRows && n.rfx && r.sheetRows <= n.rfx.r) break;
				if (!n.t) n.t = "";
				delete n.rfx;
				t.push(n);
				break;
			case 3072:
				break;
			case 35:
				i = true;
				break;
			case 36:
				i = false;
				break;
			case 37:
				break;
			case 38:
				break;
			default:
				if (f.T) {} else if (!i || r.WTF) throw new Error("Unexpected record 0x" + l.toString(16));
			}
		});
		return t
	}
	function Gu(e) {
		var r = Sa();
		var t = [];
		xa(r, 628);
		xa(r, 630);
		e.forEach(function(e) {
			e[1].forEach(function(e) {
				if (t.indexOf(e.a) > -1) return;
				t.push(e.a.slice(0, 54));
				xa(r, 632, $u(e.a));
				if (e.T && e.ID && t.indexOf("tc=" + e.ID) == -1) {
					t.push("tc=" + e.ID);
					xa(r, 632, $u("tc=" + e.ID))
				}
			})
		});
		xa(r, 631);
		xa(r, 633);
		e.forEach(function(e) {
			e[1].forEach(function(a) {
				var n = -1;
				if (a.ID) n = t.indexOf("tc=" + a.ID);
				if (n == -1 && e[1][0].T && e[1][0].ID) n = t.indexOf("tc=" + e[1][0].ID);
				if (n == -1) n = t.indexOf(a.a);
				a.iauthor = n;
				var i = {
					s: Wa(e[0]),
					e: Wa(e[0])
				};
				xa(r, 635, Hu([i, a]));
				if (a.t && a.t.length > 0) xa(r, 637, on(a));
				xa(r, 636);
				delete a.iauthor
			})
		});
		xa(r, 634);
		xa(r, 629);
		return r.end()
	}
	var ju = "application/vnd.ms-office.vbaProject";
	function Ku(e) {
		var r = Qe.utils.cfb_new({
			root: "R"
		});
		e.FullPaths.forEach(function(t, a) {
			if (t.slice(-1) === "/" || !t.match(/_VBA_PROJECT_CUR/)) return;
			var n = t.replace(/^[^\/]*/, "R").replace(/\/_VBA_PROJECT_CUR\u0000*/, "");
			Qe.utils.cfb_add(r, n, e.FileIndex[a].content)
		});
		return Qe.write(r)
	}
	function Yu(e, r) {
		r.FullPaths.forEach(function(t, a) {
			if (a == 0) return;
			var n = t.replace(/[^\/]*[\/]/, "/_VBA_PROJECT_CUR/");
			if (n.slice(-1) !== "/") Qe.utils.cfb_add(e, n, r.FileIndex[a].content)
		})
	}
	var Zu = ["xlsb", "xlsm", "xlam", "biff8", "xla"];
	function Ju() {
		return {
			"!type": "dialog"
		}
	}
	function qu() {
		return {
			"!type": "dialog"
		}
	}
	function Qu() {
		return {
			"!type": "macro"
		}
	}
	function eh() {
		return {
			"!type": "macro"
		}
	}
	var rh = function() {
		var e = /(^|[^A-Za-z_])R(\[?-?\d+\]|[1-9]\d*|)C(\[?-?\d+\]|[1-9]\d*|)(?![A-Za-z0-9_])/g;
		var r = {
			r: 0,
			c: 0
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
			var f = a.length > 0 ? parseInt(a, 10) | 0 : 0,
			l = n.length > 0 ? parseInt(n, 10) | 0 : 0;
			if (i) l += r.c;
			else--l;
			if (s) f += r.r;
			else--f;
			return t + (i ? "": "$") + La(l) + (s ? "": "$") + Na(f)
		}
		return function a(n, i) {
			r = i;
			return n.replace(e, t)
		}
	} ();
	var th = /(^|[^._A-Z0-9])(\$?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])(\$?)(\d{1,7})(?![_.\(A-Za-z0-9])/g;
	try {
		th = /(^|[^._A-Z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)(10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})(?![_.\(A-Za-z0-9])/g
	} catch(ah) {}
	var nh = function() {
		return function e(r, t) {
			return r.replace(th,
			function(e, r, a, n, i, s) {
				var f = Pa(n) - (a ? 0 : t.c);
				var l = Ia(s) - (i ? 0 : t.r);
				var o = i == "$" ? l + 1 : l == 0 ? "": "[" + l + "]";
				var c = a == "$" ? f + 1 : f == 0 ? "": "[" + f + "]";
				return r + "R" + o + "C" + c
			})
		}
	} ();
	function ih(e, r) {
		return e.replace(th,
		function(e, t, a, n, i, s) {
			return t + (a == "$" ? a + n: La(Pa(n) + r.c)) + (i == "$" ? i + s: Na(Ia(s) + r.r))
		})
	}
	function sh(e, r, t) {
		var a = Ha(r),
		n = a.s,
		i = Wa(t);
		var s = {
			r: i.r - n.r,
			c: i.c - n.c
		};
		return ih(e, s)
	}
	function fh(e) {
		if (e.length == 1) return false;
		return true
	}
	function lh(e) {
		return e.replace(/_xlfn\./g, "")
	}
	function oh(e) {
		e.l += 1;
		return
	}
	function ch(e, r) {
		var t = e._R(r == 1 ? 1 : 2);
		return [t & 16383, t >> 14 & 1, t >> 15 & 1]
	}
	function uh(e, r, t) {
		var a = 2;
		if (t) {
			if (t.biff >= 2 && t.biff <= 5) return hh(e, r, t);
			else if (t.biff == 12) a = 4
		}
		var n = e._R(a),
		i = e._R(a);
		var s = ch(e, 2);
		var f = ch(e, 2);
		return {
			s: {
				r: n,
				c: s[0],
				cRel: s[1],
				rRel: s[2]
			},
			e: {
				r: i,
				c: f[0],
				cRel: f[1],
				rRel: f[2]
			}
		}
	}
	function hh(e) {
		var r = ch(e, 2),
		t = ch(e, 2);
		var a = e._R(1);
		var n = e._R(1);
		return {
			s: {
				r: r[0],
				c: a,
				cRel: r[1],
				rRel: r[2]
			},
			e: {
				r: t[0],
				c: n,
				cRel: t[1],
				rRel: t[2]
			}
		}
	}
	function dh(e, r, t) {
		if (t.biff < 8) return hh(e, r, t);
		var a = e._R(t.biff == 12 ? 4 : 2),
		n = e._R(t.biff == 12 ? 4 : 2);
		var i = ch(e, 2);
		var s = ch(e, 2);
		return {
			s: {
				r: a,
				c: i[0],
				cRel: i[1],
				rRel: i[2]
			},
			e: {
				r: n,
				c: s[0],
				cRel: s[1],
				rRel: s[2]
			}
		}
	}
	function vh(e, r, t) {
		if (t && t.biff >= 2 && t.biff <= 5) return ph(e, r, t);
		var a = e._R(t && t.biff == 12 ? 4 : 2);
		var n = ch(e, 2);
		return {
			r: a,
			c: n[0],
			cRel: n[1],
			rRel: n[2]
		}
	}
	function ph(e) {
		var r = ch(e, 2);
		var t = e._R(1);
		return {
			r: r[0],
			c: t,
			cRel: r[1],
			rRel: r[2]
		}
	}
	function mh(e) {
		var r = e._R(2);
		var t = e._R(2);
		return {
			r: r,
			c: t & 255,
			fQuoted: !!(t & 16384),
			cRel: t >> 15,
			rRel: t >> 15
		}
	}
	function bh(e, r, t) {
		var a = t && t.biff ? t.biff: 8;
		if (a >= 2 && a <= 5) return gh(e, r, t);
		var n = e._R(a >= 12 ? 4 : 2);
		var i = e._R(2);
		var s = (i & 16384) >> 14,
		f = (i & 32768) >> 15;
		i &= 16383;
		if (f == 1) while (n > 524287) n -= 1048576;
		if (s == 1) while (i > 8191) i = i - 16384;
		return {
			r: n,
			c: i,
			cRel: s,
			rRel: f
		}
	}
	function gh(e) {
		var r = e._R(2);
		var t = e._R(1);
		var a = (r & 32768) >> 15,
		n = (r & 16384) >> 14;
		r &= 16383;
		if (a == 1 && r >= 8192) r = r - 16384;
		if (n == 1 && t >= 128) t = t - 256;
		return {
			r: r,
			c: t,
			cRel: n,
			rRel: a
		}
	}
	function wh(e, r, t) {
		var a = (e[e.l++] & 96) >> 5;
		var n = uh(e, t.biff >= 2 && t.biff <= 5 ? 6 : 8, t);
		return [a, n]
	}
	function kh(e, r, t) {
		var a = (e[e.l++] & 96) >> 5;
		var n = e._R(2, "i");
		var i = 8;
		if (t) switch (t.biff) {
		case 5:
			e.l += 12;
			i = 6;
			break;
		case 12:
			i = 12;
			break;
		}
		var s = uh(e, i, t);
		return [a, n, s]
	}
	function Th(e, r, t) {
		var a = (e[e.l++] & 96) >> 5;
		e.l += t && t.biff > 8 ? 12 : t.biff < 8 ? 6 : 8;
		return [a]
	}
	function yh(e, r, t) {
		var a = (e[e.l++] & 96) >> 5;
		var n = e._R(2);
		var i = 8;
		if (t) switch (t.biff) {
		case 5:
			e.l += 12;
			i = 6;
			break;
		case 12:
			i = 12;
			break;
		}
		e.l += i;
		return [a, n]
	}
	function Eh(e, r, t) {
		var a = (e[e.l++] & 96) >> 5;
		var n = dh(e, r - 1, t);
		return [a, n]
	}
	function _h(e, r, t) {
		var a = (e[e.l++] & 96) >> 5;
		e.l += t.biff == 2 ? 6 : t.biff == 12 ? 14 : 7;
		return [a]
	}
	function Sh(e) {
		var r = e[e.l + 1] & 1;
		var t = 1;
		e.l += 4;
		return [r, t]
	}
	function xh(e, r, t) {
		e.l += 2;
		var a = e._R(t && t.biff == 2 ? 1 : 2);
		var n = [];
		for (var i = 0; i <= a; ++i) n.push(e._R(t && t.biff == 2 ? 1 : 2));
		return n
	}
	function Ah(e, r, t) {
		var a = e[e.l + 1] & 255 ? 1 : 0;
		e.l += 2;
		return [a, e._R(t && t.biff == 2 ? 1 : 2)]
	}
	function Ch(e, r, t) {
		var a = e[e.l + 1] & 255 ? 1 : 0;
		e.l += 2;
		return [a, e._R(t && t.biff == 2 ? 1 : 2)]
	}
	function Rh(e) {
		var r = e[e.l + 1] & 255 ? 1 : 0;
		e.l += 2;
		return [r, e._R(2)]
	}
	function Oh(e, r, t) {
		var a = e[e.l + 1] & 255 ? 1 : 0;
		e.l += t && t.biff == 2 ? 3 : 4;
		return [a]
	}
	function Ih(e) {
		var r = e._R(1),
		t = e._R(1);
		return [r, t]
	}
	function Nh(e) {
		e._R(2);
		return Ih(e, 2)
	}
	function Fh(e) {
		e._R(2);
		return Ih(e, 2)
	}
	function Dh(e, r, t) {
		var a = (e[e.l] & 96) >> 5;
		e.l += 1;
		var n = vh(e, 0, t);
		return [a, n]
	}
	function Ph(e, r, t) {
		var a = (e[e.l] & 96) >> 5;
		e.l += 1;
		var n = bh(e, 0, t);
		return [a, n]
	}
	function Lh(e, r, t) {
		var a = (e[e.l] & 96) >> 5;
		e.l += 1;
		var n = e._R(2);
		if (t && t.biff == 5) e.l += 12;
		var i = vh(e, 0, t);
		return [a, n, i]
	}
	function Mh(e, r, t) {
		var a = (e[e.l] & 96) >> 5;
		e.l += 1;
		var n = e._R(t && t.biff <= 3 ? 1 : 2);
		return [uv[n], cv[n], a]
	}
	function Uh(e, r, t) {
		var a = e[e.l++];
		var n = e._R(1),
		i = t && t.biff <= 3 ? [a == 88 ? -1 : 0, e._R(1)] : Bh(e);
		return [n, (i[0] === 0 ? cv: ov)[i[1]]]
	}
	function Bh(e) {
		return [e[e.l + 1] >> 7, e._R(2) & 32767]
	}
	function Wh(e, r, t) {
		e.l += t && t.biff == 2 ? 3 : 4;
		return
	}
	function zh(e, r, t) {
		e.l++;
		if (t && t.biff == 12) return [e._R(4, "i"), 0];
		var a = e._R(2);
		var n = e._R(t && t.biff == 2 ? 1 : 2);
		return [a, n]
	}
	function Hh(e) {
		e.l++;
		return ti[e._R(1)]
	}
	function Vh(e) {
		e.l++;
		return e._R(2)
	}
	function $h(e) {
		e.l++;
		return e._R(1) !== 0
	}
	function Xh(e) {
		e.l++;
		return An(e, 8)
	}
	function Gh(e, r, t) {
		e.l++;
		return gs(e, r - 1, t)
	}
	function jh(e, r) {
		var t = [e._R(1)];
		if (r == 12) switch (t[0]) {
		case 2:
			t[0] = 4;
			break;
		case 4:
			t[0] = 16;
			break;
		case 0:
			t[0] = 1;
			break;
		case 1:
			t[0] = 2;
			break;
		}
		switch (t[0]) {
		case 4:
			t[1] = us(e, 1) ? "TRUE": "FALSE";
			if (r != 12) e.l += 7;
			break;
		case 37:
			;
		case 16:
			t[1] = ti[e[e.l]];
			e.l += r == 12 ? 4 : 8;
			break;
		case 0:
			e.l += 8;
			break;
		case 1:
			t[1] = An(e, 8);
			break;
		case 2:
			t[1] = Es(e, 0, {
				biff: r > 0 && r < 8 ? 2 : r
			});
			break;
		default:
			throw new Error("Bad SerAr: " + t[0]);
		}
		return t
	}
	function Kh(e, r, t) {
		var a = e._R(t.biff == 12 ? 4 : 2);
		var n = [];
		for (var i = 0; i != a; ++i) n.push((t.biff == 12 ? Sn: Hs)(e, 8));
		return n
	}
	function Yh(e, r, t) {
		var a = 0,
		n = 0;
		if (t.biff == 12) {
			a = e._R(4);
			n = e._R(4)
		} else {
			n = 1 + e._R(1);
			a = 1 + e._R(2)
		}
		if (t.biff >= 2 && t.biff < 8) {--a;
			if (--n == 0) n = 256
		}
		for (var i = 0,
		s = []; i != a && (s[i] = []); ++i) for (var f = 0; f != n; ++f) s[i][f] = jh(e, t.biff);
		return s
	}
	function Zh(e, r, t) {
		var a = e._R(1) >>> 5 & 3;
		var n = !t || t.biff >= 8 ? 4 : 2;
		var i = e._R(n);
		switch (t.biff) {
		case 2:
			e.l += 5;
			break;
		case 3:
			;
		case 4:
			e.l += 8;
			break;
		case 5:
			e.l += 12;
			break;
		}
		return [a, 0, i]
	}
	function Jh(e, r, t) {
		if (t.biff == 5) return qh(e, r, t);
		var a = e._R(1) >>> 5 & 3;
		var n = e._R(2);
		var i = e._R(4);
		return [a, n, i]
	}
	function qh(e) {
		var r = e._R(1) >>> 5 & 3;
		var t = e._R(2, "i");
		e.l += 8;
		var a = e._R(2);
		e.l += 12;
		return [r, t, a]
	}
	function Qh(e, r, t) {
		var a = e._R(1) >>> 5 & 3;
		e.l += t && t.biff == 2 ? 3 : 4;
		var n = e._R(t && t.biff == 2 ? 1 : 2);
		return [a, n]
	}
	function ed(e, r, t) {
		var a = e._R(1) >>> 5 & 3;
		var n = e._R(t && t.biff == 2 ? 1 : 2);
		return [a, n]
	}
	function rd(e, r, t) {
		var a = e._R(1) >>> 5 & 3;
		e.l += 4;
		if (t.biff < 8) e.l--;
		if (t.biff == 12) e.l += 2;
		return [a]
	}
	function td(e, r, t) {
		var a = (e[e.l++] & 96) >> 5;
		var n = e._R(2);
		var i = 4;
		if (t) switch (t.biff) {
		case 5:
			i = 15;
			break;
		case 12:
			i = 6;
			break;
		}
		e.l += i;
		return [a, n]
	}
	var ad = ya;
	var nd = ya;
	var id = ya;
	function sd(e, r, t) {
		e.l += 2;
		return [mh(e, 4, t)]
	}
	function fd(e) {
		e.l += 6;
		return []
	}
	var ld = sd;
	var od = fd;
	var cd = fd;
	var ud = sd;
	function hd(e) {
		e.l += 2;
		return [ds(e), e._R(2) & 1]
	}
	var dd = sd;
	var vd = hd;
	var pd = fd;
	var md = sd;
	var bd = sd;
	var gd = ["Data", "All", "Headers", "??", "?Data2", "??", "?DataHeaders", "??", "Totals", "??", "??", "??", "?DataTotals", "??", "??", "??", "?Current"];
	function wd(e) {
		e.l += 2;
		var r = e._R(2);
		var t = e._R(2);
		var a = e._R(4);
		var n = e._R(2);
		var i = e._R(2);
		var s = gd[t >> 2 & 31];
		return {
			ixti: r,
			coltype: t & 3,
			rt: s,
			idx: a,
			c: n,
			C: i
		}
	}
	function kd(e) {
		e.l += 2;
		return [e._R(4)]
	}
	function Td(e, r, t) {
		e.l += 5;
		e.l += 2;
		e.l += t.biff == 2 ? 1 : 4;
		return ["PTGSHEET"]
	}
	function yd(e, r, t) {
		e.l += t.biff == 2 ? 4 : 5;
		return ["PTGENDSHEET"]
	}
	function Ed(e) {
		var r = e._R(1) >>> 5 & 3;
		var t = e._R(2);
		return [r, t]
	}
	function _d(e) {
		var r = e._R(1) >>> 5 & 3;
		var t = e._R(2);
		return [r, t]
	}
	function Sd(e) {
		e.l += 4;
		return [0, 0]
	}
	var xd = {
		1 : {
			n: "PtgExp",
			f: zh
		},
		2 : {
			n: "PtgTbl",
			f: id
		},
		3 : {
			n: "PtgAdd",
			f: oh
		},
		4 : {
			n: "PtgSub",
			f: oh
		},
		5 : {
			n: "PtgMul",
			f: oh
		},
		6 : {
			n: "PtgDiv",
			f: oh
		},
		7 : {
			n: "PtgPower",
			f: oh
		},
		8 : {
			n: "PtgConcat",
			f: oh
		},
		9 : {
			n: "PtgLt",
			f: oh
		},
		10 : {
			n: "PtgLe",
			f: oh
		},
		11 : {
			n: "PtgEq",
			f: oh
		},
		12 : {
			n: "PtgGe",
			f: oh
		},
		13 : {
			n: "PtgGt",
			f: oh
		},
		14 : {
			n: "PtgNe",
			f: oh
		},
		15 : {
			n: "PtgIsect",
			f: oh
		},
		16 : {
			n: "PtgUnion",
			f: oh
		},
		17 : {
			n: "PtgRange",
			f: oh
		},
		18 : {
			n: "PtgUplus",
			f: oh
		},
		19 : {
			n: "PtgUminus",
			f: oh
		},
		20 : {
			n: "PtgPercent",
			f: oh
		},
		21 : {
			n: "PtgParen",
			f: oh
		},
		22 : {
			n: "PtgMissArg",
			f: oh
		},
		23 : {
			n: "PtgStr",
			f: Gh
		},
		26 : {
			n: "PtgSheet",
			f: Td
		},
		27 : {
			n: "PtgEndSheet",
			f: yd
		},
		28 : {
			n: "PtgErr",
			f: Hh
		},
		29 : {
			n: "PtgBool",
			f: $h
		},
		30 : {
			n: "PtgInt",
			f: Vh
		},
		31 : {
			n: "PtgNum",
			f: Xh
		},
		32 : {
			n: "PtgArray",
			f: _h
		},
		33 : {
			n: "PtgFunc",
			f: Mh
		},
		34 : {
			n: "PtgFuncVar",
			f: Uh
		},
		35 : {
			n: "PtgName",
			f: Zh
		},
		36 : {
			n: "PtgRef",
			f: Dh
		},
		37 : {
			n: "PtgArea",
			f: wh
		},
		38 : {
			n: "PtgMemArea",
			f: Qh
		},
		39 : {
			n: "PtgMemErr",
			f: ad
		},
		40 : {
			n: "PtgMemNoMem",
			f: nd
		},
		41 : {
			n: "PtgMemFunc",
			f: ed
		},
		42 : {
			n: "PtgRefErr",
			f: rd
		},
		43 : {
			n: "PtgAreaErr",
			f: Th
		},
		44 : {
			n: "PtgRefN",
			f: Ph
		},
		45 : {
			n: "PtgAreaN",
			f: Eh
		},
		46 : {
			n: "PtgMemAreaN",
			f: Ed
		},
		47 : {
			n: "PtgMemNoMemN",
			f: _d
		},
		57 : {
			n: "PtgNameX",
			f: Jh
		},
		58 : {
			n: "PtgRef3d",
			f: Lh
		},
		59 : {
			n: "PtgArea3d",
			f: kh
		},
		60 : {
			n: "PtgRefErr3d",
			f: td
		},
		61 : {
			n: "PtgAreaErr3d",
			f: yh
		},
		255 : {}
	};
	var Ad = {
		64 : 32,
		96 : 32,
		65 : 33,
		97 : 33,
		66 : 34,
		98 : 34,
		67 : 35,
		99 : 35,
		68 : 36,
		100 : 36,
		69 : 37,
		101 : 37,
		70 : 38,
		102 : 38,
		71 : 39,
		103 : 39,
		72 : 40,
		104 : 40,
		73 : 41,
		105 : 41,
		74 : 42,
		106 : 42,
		75 : 43,
		107 : 43,
		76 : 44,
		108 : 44,
		77 : 45,
		109 : 45,
		78 : 46,
		110 : 46,
		79 : 47,
		111 : 47,
		88 : 34,
		120 : 34,
		89 : 57,
		121 : 57,
		90 : 58,
		122 : 58,
		91 : 59,
		123 : 59,
		92 : 60,
		124 : 60,
		93 : 61,
		125 : 61
	};
	var Cd = {
		1 : {
			n: "PtgElfLel",
			f: hd
		},
		2 : {
			n: "PtgElfRw",
			f: md
		},
		3 : {
			n: "PtgElfCol",
			f: ld
		},
		6 : {
			n: "PtgElfRwV",
			f: bd
		},
		7 : {
			n: "PtgElfColV",
			f: ud
		},
		10 : {
			n: "PtgElfRadical",
			f: dd
		},
		11 : {
			n: "PtgElfRadicalS",
			f: pd
		},
		13 : {
			n: "PtgElfColS",
			f: od
		},
		15 : {
			n: "PtgElfColSV",
			f: cd
		},
		16 : {
			n: "PtgElfRadicalLel",
			f: vd
		},
		25 : {
			n: "PtgList",
			f: wd
		},
		29 : {
			n: "PtgSxName",
			f: kd
		},
		255 : {}
	};
	var Rd = {
		0 : {
			n: "PtgAttrNoop",
			f: Sd
		},
		1 : {
			n: "PtgAttrSemi",
			f: Oh
		},
		2 : {
			n: "PtgAttrIf",
			f: Ch
		},
		4 : {
			n: "PtgAttrChoose",
			f: xh
		},
		8 : {
			n: "PtgAttrGoto",
			f: Ah
		},
		16 : {
			n: "PtgAttrSum",
			f: Wh
		},
		32 : {
			n: "PtgAttrBaxcel",
			f: Sh
		},
		33 : {
			n: "PtgAttrBaxcel",
			f: Sh
		},
		64 : {
			n: "PtgAttrSpace",
			f: Nh
		},
		65 : {
			n: "PtgAttrSpaceSemi",
			f: Fh
		},
		128 : {
			n: "PtgAttrIfError",
			f: Rh
		},
		255 : {}
	};
	function Od(e, r, t, a) {
		if (a.biff < 8) return ya(e, r);
		var n = e.l + r;
		var i = [];
		for (var s = 0; s !== t.length; ++s) {
			switch (t[s][0]) {
			case "PtgArray":
				t[s][1] = Yh(e, 0, a);
				i.push(t[s][1]);
				break;
			case "PtgMemArea":
				t[s][2] = Kh(e, t[s][1], a);
				i.push(t[s][2]);
				break;
			case "PtgExp":
				if (a && a.biff == 12) {
					t[s][1][1] = e._R(4);
					i.push(t[s][1])
				}
				break;
			case "PtgList":
				;
			case "PtgElfRadicalS":
				;
			case "PtgElfColS":
				;
			case "PtgElfColSV":
				throw "Unsupported " + t[s][0];
			default:
				break;
			}
		}
		r = n - e.l;
		if (r !== 0) i.push(ya(e, r));
		return i
	}
	function Id(e, r, t) {
		var a = e.l + r;
		var n, i, s = [];
		while (a != e.l) {
			r = a - e.l;
			i = e[e.l];
			n = xd[i] || xd[Ad[i]];
			if (i === 24 || i === 25) n = (i === 24 ? Cd: Rd)[e[e.l + 1]];
			if (!n || !n.f) {
				ya(e, r)
			} else {
				s.push([n.n, n.f(e, r, t)])
			}
		}
		return s
	}
	function Nd(e) {
		var r = [];
		for (var t = 0; t < e.length; ++t) {
			var a = e[t],
			n = [];
			for (var i = 0; i < a.length; ++i) {
				var s = a[i];
				if (s) switch (s[0]) {
				case 2:
					n.push('"' + s[1].replace(/"/g, '""') + '"');
					break;
				default:
					n.push(s[1]);
				} else n.push("")
			}
			r.push(n.join(","))
		}
		return r.join(";")
	}
	var Fd = {
		PtgAdd: "+",
		PtgConcat: "&",
		PtgDiv: "/",
		PtgEq: "=",
		PtgGe: ">=",
		PtgGt: ">",
		PtgLe: "<=",
		PtgLt: "<",
		PtgMul: "*",
		PtgNe: "<>",
		PtgPower: "^",
		PtgSub: "-"
	};
	function Dd(e, r) {
		var t = e.lastIndexOf("!"),
		a = r.lastIndexOf("!");
		if (t == -1 && a == -1) return e + ":" + r;
		if (t > 0 && a > 0 && e.slice(0, t).toLowerCase() == r.slice(0, a).toLowerCase()) return e + ":" + r.slice(a + 1);
		console.error("Cannot hydrate range", e, r);
		return e + ":" + r
	}
	function Pd(e, r, t) {
		if (!e) return "SH33TJSERR0";
		if (t.biff > 8 && (!e.XTI || !e.XTI[r])) return e.SheetNames[r];
		if (!e.XTI) return "SH33TJSERR6";
		var a = e.XTI[r];
		if (t.biff < 8) {
			if (r > 1e4) r -= 65536;
			if (r < 0) r = -r;
			return r == 0 ? "": e.XTI[r - 1]
		}
		if (!a) return "SH33TJSERR1";
		var n = "";
		if (t.biff > 8) switch (e[a[0]][0]) {
		case 357:
			n = a[1] == -1 ? "#REF": e.SheetNames[a[1]];
			return a[1] == a[2] ? n: n + ":" + e.SheetNames[a[2]];
		case 358:
			if (t.SID != null) return e.SheetNames[t.SID];
			return "SH33TJSSAME" + e[a[0]][0];
		case 355:
			;
		default:
			return "SH33TJSSRC" + e[a[0]][0];
		}
		switch (e[a[0]][0][0]) {
		case 1025:
			n = a[1] == -1 ? "#REF": e.SheetNames[a[1]] || "SH33TJSERR3";
			return a[1] == a[2] ? n: n + ":" + e.SheetNames[a[2]];
		case 14849:
			return e[a[0]].slice(1).map(function(e) {
				return e.Name
			}).join(";;");
		default:
			if (!e[a[0]][0][3]) return "SH33TJSERR2";
			n = a[1] == -1 ? "#REF": e[a[0]][0][3][a[1]] || "SH33TJSERR4";
			return a[1] == a[2] ? n: n + ":" + e[a[0]][0][3][a[2]];
		}
	}
	function Ld(e, r, t) {
		var a = Pd(e, r, t);
		return a == "#REF" ? a: Xa(a, t)
	}
	function Md(e, r, t, a, n) {
		var i = n && n.biff || 8;
		var s = {
			s: {
				c: 0,
				r: 0
			},
			e: {
				c: 0,
				r: 0
			}
		};
		var f = [],
		l,
		o,
		c,
		u = 0,
		h = 0,
		d,
		v = "";
		if (!e[0] || !e[0][0]) return "";
		var p = -1,
		m = "";
		for (var b = 0,
		g = e[0].length; b < g; ++b) {
			var w = e[0][b];
			switch (w[0]) {
			case "PtgUminus":
				f.push("-" + f.pop());
				break;
			case "PtgUplus":
				f.push("+" + f.pop());
				break;
			case "PtgPercent":
				f.push(f.pop() + "%");
				break;
			case "PtgAdd":
				;
			case "PtgConcat":
				;
			case "PtgDiv":
				;
			case "PtgEq":
				;
			case "PtgGe":
				;
			case "PtgGt":
				;
			case "PtgLe":
				;
			case "PtgLt":
				;
			case "PtgMul":
				;
			case "PtgNe":
				;
			case "PtgPower":
				;
			case "PtgSub":
				l = f.pop();
				o = f.pop();
				if (p >= 0) {
					switch (e[0][p][1][0]) {
					case 0:
						m = yr(" ", e[0][p][1][1]);
						break;
					case 1:
						m = yr("\r", e[0][p][1][1]);
						break;
					default:
						m = "";
						if (n.WTF) throw new Error("Unexpected PtgAttrSpaceType " + e[0][p][1][0]);
					}
					o = o + m;
					p = -1
				}
				f.push(o + Fd[w[0]] + l);
				break;
			case "PtgIsect":
				l = f.pop();
				o = f.pop();
				f.push(o + " " + l);
				break;
			case "PtgUnion":
				l = f.pop();
				o = f.pop();
				f.push(o + "," + l);
				break;
			case "PtgRange":
				l = f.pop();
				o = f.pop();
				f.push(Dd(o, l));
				break;
			case "PtgAttrChoose":
				break;
			case "PtgAttrGoto":
				break;
			case "PtgAttrIf":
				break;
			case "PtgAttrIfError":
				break;
			case "PtgRef":
				c = Aa(w[1][1], s, n);
				f.push(Ra(c, i));
				break;
			case "PtgRefN":
				c = t ? Aa(w[1][1], t, n) : w[1][1];
				f.push(Ra(c, i));
				break;
			case "PtgRef3d":
				u = w[1][1];
				c = Aa(w[1][2], s, n);
				v = Ld(a, u, n);
				var k = v;
				f.push(v + "!" + Ra(c, i));
				break;
			case "PtgFunc":
				;
			case "PtgFuncVar":
				var T = w[1][0],
				y = w[1][1];
				if (!T) T = 0;
				T &= 127;
				var E = T == 0 ? [] : f.slice(-T);
				f.length -= T;
				if (y === "User") y = E.shift();
				f.push(y + "(" + E.join(",") + ")");
				break;
			case "PtgBool":
				f.push(w[1] ? "TRUE": "FALSE");
				break;
			case "PtgInt":
				f.push(w[1]);
				break;
			case "PtgNum":
				f.push(String(w[1]));
				break;
			case "PtgStr":
				f.push('"' + w[1].replace(/"/g, '""') + '"');
				break;
			case "PtgErr":
				f.push(w[1]);
				break;
			case "PtgAreaN":
				d = Ca(w[1][1], t ? {
					s: t
				}: s, n);
				f.push(Oa(d, n));
				break;
			case "PtgArea":
				d = Ca(w[1][1], s, n);
				f.push(Oa(d, n));
				break;
			case "PtgArea3d":
				u = w[1][1];
				d = w[1][2];
				v = Ld(a, u, n);
				f.push(v + "!" + Oa(d, n));
				break;
			case "PtgAttrSum":
				f.push("SUM(" + f.pop() + ")");
				break;
			case "PtgAttrBaxcel":
				;
			case "PtgAttrSemi":
				break;
			case "PtgName":
				h = w[1][2];
				var _ = (a.names || [])[h - 1] || (a[0] || [])[h];
				var S = _ ? _.Name: "SH33TJSNAME" + String(h);
				if (S && S.slice(0, 6) == "_xlfn." && !n.xlfn) S = S.slice(6);
				f.push(S);
				break;
			case "PtgNameX":
				var x = w[1][1];
				h = w[1][2];
				var A;
				if (n.biff <= 5) {
					if (x < 0) x = -x;
					if (a[x]) A = a[x][h]
				} else {
					var C = "";
					if (((a[x] || [])[0] || [])[0] == 14849) {} else if (((a[x] || [])[0] || [])[0] == 1025) {
						if (a[x][h] && a[x][h].itab > 0) {
							C = a.SheetNames[a[x][h].itab - 1] + "!"
						}
					} else C = a.SheetNames[h - 1] + "!";
					if (a[x] && a[x][h]) C += a[x][h].Name;
					else if (a[0] && a[0][h]) C += a[0][h].Name;
					else {
						var R = (Pd(a, x, n) || "").split(";;");
						if (R[h - 1]) C = R[h - 1];
						else C += "SH33TJSERRX"
					}
					f.push(C);
					break
				}
				if (!A) A = {
					Name: "SH33TJSERRY"
				};
				f.push(A.Name);
				break;
			case "PtgParen":
				var O = "(",
				I = ")";
				if (p >= 0) {
					m = "";
					switch (e[0][p][1][0]) {
					case 2:
						O = yr(" ", e[0][p][1][1]) + O;
						break;
					case 3:
						O = yr("\r", e[0][p][1][1]) + O;
						break;
					case 4:
						I = yr(" ", e[0][p][1][1]) + I;
						break;
					case 5:
						I = yr("\r", e[0][p][1][1]) + I;
						break;
					default:
						if (n.WTF) throw new Error("Unexpected PtgAttrSpaceType " + e[0][p][1][0]);
					}
					p = -1
				}
				f.push(O + f.pop() + I);
				break;
			case "PtgRefErr":
				f.push("#REF!");
				break;
			case "PtgRefErr3d":
				f.push("#REF!");
				break;
			case "PtgExp":
				c = {
					c: w[1][1],
					r: w[1][0]
				};
				var N = {
					c: t.c,
					r: t.r
				};
				if (a.sharedf[za(c)]) {
					var F = a.sharedf[za(c)];
					f.push(Md(F, s, N, a, n))
				} else {
					var D = false;
					for (l = 0; l != a.arrayf.length; ++l) {
						o = a.arrayf[l];
						if (c.c < o[0].s.c || c.c > o[0].e.c) continue;
						if (c.r < o[0].s.r || c.r > o[0].e.r) continue;
						f.push(Md(o[1], s, N, a, n));
						D = true;
						break
					}
					if (!D) f.push(w[1])
				}
				break;
			case "PtgArray":
				f.push("{" + Nd(w[1]) + "}");
				break;
			case "PtgMemArea":
				break;
			case "PtgAttrSpace":
				;
			case "PtgAttrSpaceSemi":
				p = b;
				break;
			case "PtgTbl":
				break;
			case "PtgMemErr":
				break;
			case "PtgMissArg":
				f.push("");
				break;
			case "PtgAreaErr":
				f.push("#REF!");
				break;
			case "PtgAreaErr3d":
				f.push("#REF!");
				break;
			case "PtgList":
				f.push("Table" + w[1].idx + "[#" + w[1].rt + "]");
				break;
			case "PtgMemAreaN":
				;
			case "PtgMemNoMemN":
				;
			case "PtgAttrNoop":
				;
			case "PtgSheet":
				;
			case "PtgEndSheet":
				break;
			case "PtgMemFunc":
				break;
			case "PtgMemNoMem":
				break;
			case "PtgElfCol":
				;
			case "PtgElfColS":
				;
			case "PtgElfColSV":
				;
			case "PtgElfColV":
				;
			case "PtgElfLel":
				;
			case "PtgElfRadical":
				;
			case "PtgElfRadicalLel":
				;
			case "PtgElfRadicalS":
				;
			case "PtgElfRw":
				;
			case "PtgElfRwV":
				throw new Error("Unsupported ELFs");
			case "PtgSxName":
				throw new Error("Unrecognized Formula Token: " + String(w));
			default:
				throw new Error("Unrecognized Formula Token: " + String(w));
			}
			var P = ["PtgAttrSpace", "PtgAttrSpaceSemi", "PtgAttrGoto"];
			if (n.biff != 3) if (p >= 0 && P.indexOf(e[0][b][0]) == -1) {
				w = e[0][p];
				var L = true;
				switch (w[1][0]) {
				case 4:
					L = false;
				case 0:
					m = yr(" ", w[1][1]);
					break;
				case 5:
					L = false;
				case 1:
					m = yr("\r", w[1][1]);
					break;
				default:
					m = "";
					if (n.WTF) throw new Error("Unexpected PtgAttrSpaceType " + w[1][0]);
				}
				f.push((L ? m: "") + f.pop() + (L ? "": m));
				p = -1
			}
		}
		if (f.length > 1 && n.WTF) throw new Error("bad formula stack");
		if (f[0] == "TRUE") return true;
		if (f[0] == "FALSE") return false;
		return f[0]
	}
	function Ud(e, r, t) {
		var a = e.l + r,
		n = t.biff == 2 ? 1 : 2;
		var i, s = e._R(n);
		if (s == 65535) return [[], ya(e, r - 2)];
		var f = Id(e, s, t);
		if (r !== s + n) i = Od(e, r - s - n, f, t);
		e.l = a;
		return [f, i]
	}
	function Bd(e, r, t) {
		var a = e.l + r,
		n = t.biff == 2 ? 1 : 2;
		var i, s = e._R(n);
		if (s == 65535) return [[], ya(e, r - 2)];
		var f = Id(e, s, t);
		if (r !== s + n) i = Od(e, r - s - n, f, t);
		e.l = a;
		return [f, i]
	}
	function Wd(e, r, t, a) {
		var n = e.l + r;
		var i = Id(e, a, t);
		var s;
		if (n !== e.l) s = Od(e, n - e.l, i, t);
		return [i, s]
	}
	function zd(e, r, t) {
		var a = e.l + r;
		var n, i = e._R(2);
		var s = Id(e, i, t);
		if (i == 65535) return [[], ya(e, r - 2)];
		if (r !== i + 2) n = Od(e, a - i - 2, s, t);
		return [s, n]
	}
	function Hd(e) {
		var r;
		if (ca(e, e.l + 6) !== 65535) return [An(e), "n"];
		switch (e[e.l]) {
		case 0:
			e.l += 8;
			return ["String", "s"];
		case 1:
			r = e[e.l + 2] === 1;
			e.l += 8;
			return [r, "b"];
		case 2:
			r = e[e.l + 2];
			e.l += 8;
			return [r, "e"];
		case 3:
			e.l += 8;
			return ["", "s"];
		}
		return []
	}
	function Vd(e) {
		if (e == null) {
			var r = Ea(8);
			r._W(1, 3);
			r._W(1, 0);
			r._W(2, 0);
			r._W(2, 0);
			r._W(2, 65535);
			return r
		} else if (typeof e == "number") return Cn(e);
		return Cn(0)
	}
	function $d(e, r, t) {
		var a = e.l + r;
		var n = Ps(e, 6, t);
		var i = Hd(e, 8);
		var s = e._R(1);
		if (t.biff != 2) {
			e._R(1);
			if (t.biff >= 5) {
				e._R(4)
			}
		}
		var f = Bd(e, a - e.l, t);
		return {
			cell: n,
			val: i[0],
			formula: f,
			shared: s >> 3 & 1,
			tt: i[1]
		}
	}
	function Xd(e, r, t, a, n) {
		var i = Ls(r, t, n);
		var s = Vd(e.v);
		var f = Ea(6);
		var l = 1 | 32;
		f._W(2, l);
		f._W(4, 0);
		var o = Ea(e.bf.length);
		for (var c = 0; c < e.bf.length; ++c) o[c] = e.bf[c];
		var u = P([i, s, f, o]);
		return u
	}
	function Gd(e, r, t) {
		var a = e._R(4);
		var n = Id(e, a, t);
		var i = e._R(4);
		var s = i > 0 ? Od(e, i, n, t) : null;
		return [n, s]
	}
	var jd = Gd;
	var Kd = Gd;
	var Yd = Gd;
	var Zd = Gd;
	function Jd(e) {
		if ((e | 0) == e && e < Math.pow(2, 16) && e >= 0) {
			var r = Ea(11);
			r._W(4, 3);
			r._W(1, 30);
			r._W(2, e);
			r._W(4, 0);
			return r
		}
		var t = Ea(17);
		t._W(4, 11);
		t._W(1, 31);
		t._W(8, e);
		t._W(4, 0);
		return t
	}
	function qd(e) {
		var r = Ea(10);
		r._W(4, 2);
		r._W(1, 28);
		r._W(1, e);
		r._W(4, 0);
		return r
	}
	function Qd(e) {
		var r = Ea(10);
		r._W(4, 2);
		r._W(1, 29);
		r._W(1, e ? 1 : 0);
		r._W(4, 0);
		return r
	}
	function ev(e) {
		var r = Ea(7);
		r._W(4, 3 + 2 * e.length);
		r._W(1, 23);
		r._W(2, e.length);
		var t = Ea(2 * e.length);
		t._W(2 * e.length, e, "utf16le");
		var a = Ea(4);
		a._W(4, 0);
		return P([r, t, a])
	}
	function rv(e) {
		var r = Wa(e);
		var t = Ea(15);
		t._W(4, 7);
		t._W(1, 4 | 1 << 5);
		t._W(4, r.r);
		t._W(2, r.c | (e.charAt(0) == "$" ? 0 : 1) << 14 | (e.match(/\$\d/) ? 0 : 1) << 15);
		t._W(4, 0);
		return t
	}
	function tv(e, r) {
		var t = e.lastIndexOf("!");
		var a = e.slice(0, t);
		e = e.slice(t + 1);
		var n = Wa(e);
		if (a.charAt(0) == "'") a = a.slice(1, -1).replace(/''/g, "'");
		var i = Ea(17);
		i._W(4, 9);
		i._W(1, 26 | 1 << 5);
		i._W(2, 2 + r.SheetNames.map(function(e) {
			return e.toLowerCase()
		}).indexOf(a.toLowerCase()));
		i._W(4, n.r);
		i._W(2, n.c | (e.charAt(0) == "$" ? 0 : 1) << 14 | (e.match(/\$\d/) ? 0 : 1) << 15);
		i._W(4, 0);
		return i
	}
	function av(e, r) {
		var t = e.lastIndexOf("!");
		var a = e.slice(0, t);
		e = e.slice(t + 1);
		if (a.charAt(0) == "'") a = a.slice(1, -1).replace(/''/g, "'");
		var n = Ea(17);
		n._W(4, 9);
		n._W(1, 28 | 1 << 5);
		n._W(2, 2 + r.SheetNames.map(function(e) {
			return e.toLowerCase()
		}).indexOf(a.toLowerCase()));
		n._W(4, 0);
		n._W(2, 0);
		n._W(4, 0);
		return n
	}
	function nv(e) {
		var r = e.split(":"),
		t = r[0];
		var a = Ea(23);
		a._W(4, 15);
		t = r[0];
		var n = Wa(t);
		a._W(1, 4 | 1 << 5);
		a._W(4, n.r);
		a._W(2, n.c | (t.charAt(0) == "$" ? 0 : 1) << 14 | (t.match(/\$\d/) ? 0 : 1) << 15);
		a._W(4, 0);
		t = r[1];
		n = Wa(t);
		a._W(1, 4 | 1 << 5);
		a._W(4, n.r);
		a._W(2, n.c | (t.charAt(0) == "$" ? 0 : 1) << 14 | (t.match(/\$\d/) ? 0 : 1) << 15);
		a._W(4, 0);
		a._W(1, 17);
		a._W(4, 0);
		return a
	}
	function iv(e, r) {
		var t = e.lastIndexOf("!");
		var a = e.slice(0, t);
		e = e.slice(t + 1);
		if (a.charAt(0) == "'") a = a.slice(1, -1).replace(/''/g, "'");
		var n = e.split(":");
		var i = Ea(27);
		i._W(4, 19);
		var s = n[0],
		f = Wa(s);
		i._W(1, 26 | 1 << 5);
		i._W(2, 2 + r.SheetNames.map(function(e) {
			return e.toLowerCase()
		}).indexOf(a.toLowerCase()));
		i._W(4, f.r);
		i._W(2, f.c | (s.charAt(0) == "$" ? 0 : 1) << 14 | (s.match(/\$\d/) ? 0 : 1) << 15);
		s = n[1];
		f = Wa(s);
		i._W(1, 26 | 1 << 5);
		i._W(2, 2 + r.SheetNames.map(function(e) {
			return e.toLowerCase()
		}).indexOf(a.toLowerCase()));
		i._W(4, f.r);
		i._W(2, f.c | (s.charAt(0) == "$" ? 0 : 1) << 14 | (s.match(/\$\d/) ? 0 : 1) << 15);
		i._W(1, 17);
		i._W(4, 0);
		return i
	}
	function sv(e, r) {
		var t = e.lastIndexOf("!");
		var a = e.slice(0, t);
		e = e.slice(t + 1);
		if (a.charAt(0) == "'") a = a.slice(1, -1).replace(/''/g, "'");
		var n = Ha(e);
		var i = Ea(23);
		i._W(4, 15);
		i._W(1, 27 | 1 << 5);
		i._W(2, 2 + r.SheetNames.map(function(e) {
			return e.toLowerCase()
		}).indexOf(a.toLowerCase()));
		i._W(4, n.s.r);
		i._W(4, n.e.r);
		i._W(2, n.s.c);
		i._W(2, n.e.c);
		i._W(4, 0);
		return i
	}
	function fv(e, r) {
		if (typeof e == "number") return Jd(e);
		if (typeof e == "boolean") return Qd(e);
		if (/^#(DIV\/0!|GETTING_DATA|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!)$/.test(e)) return qd( + ai[e]);
		if (e.match(/^\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})$/)) return rv(e);
		if (e.match(/^\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5}):\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})$/)) return nv(e);
		if (e.match(/^#REF!\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5}):\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})$/)) return sv(e, r);
		if (e.match(/^(?:'[^\\\/?*\[\]:]*'|[^'][^\\\/?*\[\]:'`~!@#$%^()\-=+{}|;,<.>]*)!\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})$/)) return tv(e, r);
		if (e.match(/^(?:'[^\\\/?*\[\]:]*'|[^'][^\\\/?*\[\]:'`~!@#$%^()\-=+{}|;,<.>]*)!\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5}):\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})$/)) return iv(e, r);
		if (/^(?:'[^\\\/?*\[\]:]*'|[^'][^\\\/?*\[\]:'`~!@#$%^()\-=+{}|;,<.>]*)!#REF!$/.test(e)) return av(e, r);
		if (/^".*"$/.test(e)) return ev(e);
		if (/^[+-]\d+$/.test(e)) return Jd(parseInt(e, 10));
		throw "Formula |" + e + "| not supported for XLSB"
	}
	var lv = fv;
	var ov = {
		0 : "BEEP",
		1 : "OPEN",
		2 : "OPEN.LINKS",
		3 : "CLOSE.ALL",
		4 : "SAVE",
		5 : "SAVE.AS",
		6 : "FILE.DELETE",
		7 : "PAGE.SETUP",
		8 : "PRINT",
		9 : "PRINTER.SETUP",
		10 : "QUIT",
		11 : "NEW.WINDOW",
		12 : "ARRANGE.ALL",
		13 : "WINDOW.SIZE",
		14 : "WINDOW.MOVE",
		15 : "FULL",
		16 : "CLOSE",
		17 : "RUN",
		22 : "SET.PRINT.AREA",
		23 : "SET.PRINT.TITLES",
		24 : "SET.PAGE.BREAK",
		25 : "REMOVE.PAGE.BREAK",
		26 : "FONT",
		27 : "DISPLAY",
		28 : "PROTECT.DOCUMENT",
		29 : "PRECISION",
		30 : "A1.R1C1",
		31 : "CALCULATE.NOW",
		32 : "CALCULATION",
		34 : "DATA.FIND",
		35 : "EXTRACT",
		36 : "DATA.DELETE",
		37 : "SET.DATABASE",
		38 : "SET.CRITERIA",
		39 : "SORT",
		40 : "DATA.SERIES",
		41 : "TABLE",
		42 : "FORMAT.NUMBER",
		43 : "ALIGNMENT",
		44 : "STYLE",
		45 : "BORDER",
		46 : "CELL.PROTECTION",
		47 : "COLUMN.WIDTH",
		48 : "UNDO",
		49 : "CUT",
		50 : "COPY",
		51 : "PASTE",
		52 : "CLEAR",
		53 : "PASTE.SPECIAL",
		54 : "EDIT.DELETE",
		55 : "INSERT",
		56 : "FILL.RIGHT",
		57 : "FILL.DOWN",
		61 : "DEFINE.NAME",
		62 : "CREATE.NAMES",
		63 : "FORMULA.GOTO",
		64 : "FORMULA.FIND",
		65 : "SELECT.LAST.CELL",
		66 : "SHOW.ACTIVE.CELL",
		67 : "GALLERY.AREA",
		68 : "GALLERY.BAR",
		69 : "GALLERY.COLUMN",
		70 : "GALLERY.LINE",
		71 : "GALLERY.PIE",
		72 : "GALLERY.SCATTER",
		73 : "COMBINATION",
		74 : "PREFERRED",
		75 : "ADD.OVERLAY",
		76 : "GRIDLINES",
		77 : "SET.PREFERRED",
		78 : "AXES",
		79 : "LEGEND",
		80 : "ATTACH.TEXT",
		81 : "ADD.ARROW",
		82 : "SELECT.CHART",
		83 : "SELECT.PLOT.AREA",
		84 : "PATTERNS",
		85 : "MAIN.CHART",
		86 : "OVERLAY",
		87 : "SCALE",
		88 : "FORMAT.LEGEND",
		89 : "FORMAT.TEXT",
		90 : "EDIT.REPEAT",
		91 : "PARSE",
		92 : "JUSTIFY",
		93 : "HIDE",
		94 : "UNHIDE",
		95 : "WORKSPACE",
		96 : "FORMULA",
		97 : "FORMULA.FILL",
		98 : "FORMULA.ARRAY",
		99 : "DATA.FIND.NEXT",
		100 : "DATA.FIND.PREV",
		101 : "FORMULA.FIND.NEXT",
		102 : "FORMULA.FIND.PREV",
		103 : "ACTIVATE",
		104 : "ACTIVATE.NEXT",
		105 : "ACTIVATE.PREV",
		106 : "UNLOCKED.NEXT",
		107 : "UNLOCKED.PREV",
		108 : "COPY.PICTURE",
		109 : "SELECT",
		110 : "DELETE.NAME",
		111 : "DELETE.FORMAT",
		112 : "VLINE",
		113 : "HLINE",
		114 : "VPAGE",
		115 : "HPAGE",
		116 : "VSCROLL",
		117 : "HSCROLL",
		118 : "ALERT",
		119 : "NEW",
		120 : "CANCEL.COPY",
		121 : "SHOW.CLIPBOARD",
		122 : "MESSAGE",
		124 : "PASTE.LINK",
		125 : "APP.ACTIVATE",
		126 : "DELETE.ARROW",
		127 : "ROW.HEIGHT",
		128 : "FORMAT.MOVE",
		129 : "FORMAT.SIZE",
		130 : "FORMULA.REPLACE",
		131 : "SEND.KEYS",
		132 : "SELECT.SPECIAL",
		133 : "APPLY.NAMES",
		134 : "REPLACE.FONT",
		135 : "FREEZE.PANES",
		136 : "SHOW.INFO",
		137 : "SPLIT",
		138 : "ON.WINDOW",
		139 : "ON.DATA",
		140 : "DISABLE.INPUT",
		142 : "OUTLINE",
		143 : "LIST.NAMES",
		144 : "FILE.CLOSE",
		145 : "SAVE.WORKBOOK",
		146 : "DATA.FORM",
		147 : "COPY.CHART",
		148 : "ON.TIME",
		149 : "WAIT",
		150 : "FORMAT.FONT",
		151 : "FILL.UP",
		152 : "FILL.LEFT",
		153 : "DELETE.OVERLAY",
		155 : "SHORT.MENUS",
		159 : "SET.UPDATE.STATUS",
		161 : "COLOR.PALETTE",
		162 : "DELETE.STYLE",
		163 : "WINDOW.RESTORE",
		164 : "WINDOW.MAXIMIZE",
		166 : "CHANGE.LINK",
		167 : "CALCULATE.DOCUMENT",
		168 : "ON.KEY",
		169 : "APP.RESTORE",
		170 : "APP.MOVE",
		171 : "APP.SIZE",
		172 : "APP.MINIMIZE",
		173 : "APP.MAXIMIZE",
		174 : "BRING.TO.FRONT",
		175 : "SEND.TO.BACK",
		185 : "MAIN.CHART.TYPE",
		186 : "OVERLAY.CHART.TYPE",
		187 : "SELECT.END",
		188 : "OPEN.MAIL",
		189 : "SEND.MAIL",
		190 : "STANDARD.FONT",
		191 : "CONSOLIDATE",
		192 : "SORT.SPECIAL",
		193 : "GALLERY.3D.AREA",
		194 : "GALLERY.3D.COLUMN",
		195 : "GALLERY.3D.LINE",
		196 : "GALLERY.3D.PIE",
		197 : "VIEW.3D",
		198 : "GOAL.SEEK",
		199 : "WORKGROUP",
		200 : "FILL.GROUP",
		201 : "UPDATE.LINK",
		202 : "PROMOTE",
		203 : "DEMOTE",
		204 : "SHOW.DETAIL",
		206 : "UNGROUP",
		207 : "OBJECT.PROPERTIES",
		208 : "SAVE.NEW.OBJECT",
		209 : "SHARE",
		210 : "SHARE.NAME",
		211 : "DUPLICATE",
		212 : "APPLY.STYLE",
		213 : "ASSIGN.TO.OBJECT",
		214 : "OBJECT.PROTECTION",
		215 : "HIDE.OBJECT",
		216 : "SET.EXTRACT",
		217 : "CREATE.PUBLISHER",
		218 : "SUBSCRIBE.TO",
		219 : "ATTRIBUTES",
		220 : "SHOW.TOOLBAR",
		222 : "PRINT.PREVIEW",
		223 : "EDIT.COLOR",
		224 : "SHOW.LEVELS",
		225 : "FORMAT.MAIN",
		226 : "FORMAT.OVERLAY",
		227 : "ON.RECALC",
		228 : "EDIT.SERIES",
		229 : "DEFINE.STYLE",
		240 : "LINE.PRINT",
		243 : "ENTER.DATA",
		249 : "GALLERY.RADAR",
		250 : "MERGE.STYLES",
		251 : "EDITION.OPTIONS",
		252 : "PASTE.PICTURE",
		253 : "PASTE.PICTURE.LINK",
		254 : "SPELLING",
		256 : "ZOOM",
		259 : "INSERT.OBJECT",
		260 : "WINDOW.MINIMIZE",
		265 : "SOUND.NOTE",
		266 : "SOUND.PLAY",
		267 : "FORMAT.SHAPE",
		268 : "EXTEND.POLYGON",
		269 : "FORMAT.AUTO",
		272 : "GALLERY.3D.BAR",
		273 : "GALLERY.3D.SURFACE",
		274 : "FILL.AUTO",
		276 : "CUSTOMIZE.TOOLBAR",
		277 : "ADD.TOOL",
		278 : "EDIT.OBJECT",
		279 : "ON.DOUBLECLICK",
		280 : "ON.ENTRY",
		281 : "WORKBOOK.ADD",
		282 : "WORKBOOK.MOVE",
		283 : "WORKBOOK.COPY",
		284 : "WORKBOOK.OPTIONS",
		285 : "SAVE.WORKSPACE",
		288 : "CHART.WIZARD",
		289 : "DELETE.TOOL",
		290 : "MOVE.TOOL",
		291 : "WORKBOOK.SELECT",
		292 : "WORKBOOK.ACTIVATE",
		293 : "ASSIGN.TO.TOOL",
		295 : "COPY.TOOL",
		296 : "RESET.TOOL",
		297 : "CONSTRAIN.NUMERIC",
		298 : "PASTE.TOOL",
		302 : "WORKBOOK.NEW",
		305 : "SCENARIO.CELLS",
		306 : "SCENARIO.DELETE",
		307 : "SCENARIO.ADD",
		308 : "SCENARIO.EDIT",
		309 : "SCENARIO.SHOW",
		310 : "SCENARIO.SHOW.NEXT",
		311 : "SCENARIO.SUMMARY",
		312 : "PIVOT.TABLE.WIZARD",
		313 : "PIVOT.FIELD.PROPERTIES",
		314 : "PIVOT.FIELD",
		315 : "PIVOT.ITEM",
		316 : "PIVOT.ADD.FIELDS",
		318 : "OPTIONS.CALCULATION",
		319 : "OPTIONS.EDIT",
		320 : "OPTIONS.VIEW",
		321 : "ADDIN.MANAGER",
		322 : "MENU.EDITOR",
		323 : "ATTACH.TOOLBARS",
		324 : "VBAActivate",
		325 : "OPTIONS.CHART",
		328 : "VBA.INSERT.FILE",
		330 : "VBA.PROCEDURE.DEFINITION",
		336 : "ROUTING.SLIP",
		338 : "ROUTE.DOCUMENT",
		339 : "MAIL.LOGON",
		342 : "INSERT.PICTURE",
		343 : "EDIT.TOOL",
		344 : "GALLERY.DOUGHNUT",
		350 : "CHART.TREND",
		352 : "PIVOT.ITEM.PROPERTIES",
		354 : "WORKBOOK.INSERT",
		355 : "OPTIONS.TRANSITION",
		356 : "OPTIONS.GENERAL",
		370 : "FILTER.ADVANCED",
		373 : "MAIL.ADD.MAILER",
		374 : "MAIL.DELETE.MAILER",
		375 : "MAIL.REPLY",
		376 : "MAIL.REPLY.ALL",
		377 : "MAIL.FORWARD",
		378 : "MAIL.NEXT.LETTER",
		379 : "DATA.LABEL",
		380 : "INSERT.TITLE",
		381 : "FONT.PROPERTIES",
		382 : "MACRO.OPTIONS",
		383 : "WORKBOOK.HIDE",
		384 : "WORKBOOK.UNHIDE",
		385 : "WORKBOOK.DELETE",
		386 : "WORKBOOK.NAME",
		388 : "GALLERY.CUSTOM",
		390 : "ADD.CHART.AUTOFORMAT",
		391 : "DELETE.CHART.AUTOFORMAT",
		392 : "CHART.ADD.DATA",
		393 : "AUTO.OUTLINE",
		394 : "TAB.ORDER",
		395 : "SHOW.DIALOG",
		396 : "SELECT.ALL",
		397 : "UNGROUP.SHEETS",
		398 : "SUBTOTAL.CREATE",
		399 : "SUBTOTAL.REMOVE",
		400 : "RENAME.OBJECT",
		412 : "WORKBOOK.SCROLL",
		413 : "WORKBOOK.NEXT",
		414 : "WORKBOOK.PREV",
		415 : "WORKBOOK.TAB.SPLIT",
		416 : "FULL.SCREEN",
		417 : "WORKBOOK.PROTECT",
		420 : "SCROLLBAR.PROPERTIES",
		421 : "PIVOT.SHOW.PAGES",
		422 : "TEXT.TO.COLUMNS",
		423 : "FORMAT.CHARTTYPE",
		424 : "LINK.FORMAT",
		425 : "TRACER.DISPLAY",
		430 : "TRACER.NAVIGATE",
		431 : "TRACER.CLEAR",
		432 : "TRACER.ERROR",
		433 : "PIVOT.FIELD.GROUP",
		434 : "PIVOT.FIELD.UNGROUP",
		435 : "CHECKBOX.PROPERTIES",
		436 : "LABEL.PROPERTIES",
		437 : "LISTBOX.PROPERTIES",
		438 : "EDITBOX.PROPERTIES",
		439 : "PIVOT.REFRESH",
		440 : "LINK.COMBO",
		441 : "OPEN.TEXT",
		442 : "HIDE.DIALOG",
		443 : "SET.DIALOG.FOCUS",
		444 : "ENABLE.OBJECT",
		445 : "PUSHBUTTON.PROPERTIES",
		446 : "SET.DIALOG.DEFAULT",
		447 : "FILTER",
		448 : "FILTER.SHOW.ALL",
		449 : "CLEAR.OUTLINE",
		450 : "FUNCTION.WIZARD",
		451 : "ADD.LIST.ITEM",
		452 : "SET.LIST.ITEM",
		453 : "REMOVE.LIST.ITEM",
		454 : "SELECT.LIST.ITEM",
		455 : "SET.CONTROL.VALUE",
		456 : "SAVE.COPY.AS",
		458 : "OPTIONS.LISTS.ADD",
		459 : "OPTIONS.LISTS.DELETE",
		460 : "SERIES.AXES",
		461 : "SERIES.X",
		462 : "SERIES.Y",
		463 : "ERRORBAR.X",
		464 : "ERRORBAR.Y",
		465 : "FORMAT.CHART",
		466 : "SERIES.ORDER",
		467 : "MAIL.LOGOFF",
		468 : "CLEAR.ROUTING.SLIP",
		469 : "APP.ACTIVATE.MICROSOFT",
		470 : "MAIL.EDIT.MAILER",
		471 : "ON.SHEET",
		472 : "STANDARD.WIDTH",
		473 : "SCENARIO.MERGE",
		474 : "SUMMARY.INFO",
		475 : "FIND.FILE",
		476 : "ACTIVE.CELL.FONT",
		477 : "ENABLE.TIPWIZARD",
		478 : "VBA.MAKE.ADDIN",
		480 : "INSERTDATATABLE",
		481 : "WORKGROUP.OPTIONS",
		482 : "MAIL.SEND.MAILER",
		485 : "AUTOCORRECT",
		489 : "POST.DOCUMENT",
		491 : "PICKLIST",
		493 : "VIEW.SHOW",
		494 : "VIEW.DEFINE",
		495 : "VIEW.DELETE",
		509 : "SHEET.BACKGROUND",
		510 : "INSERT.MAP.OBJECT",
		511 : "OPTIONS.MENONO",
		517 : "MSOCHECKS",
		518 : "NORMAL",
		519 : "LAYOUT",
		520 : "RM.PRINT.AREA",
		521 : "CLEAR.PRINT.AREA",
		522 : "ADD.PRINT.AREA",
		523 : "MOVE.BRK",
		545 : "HIDECURR.NOTE",
		546 : "HIDEALL.NOTES",
		547 : "DELETE.NOTE",
		548 : "TRAVERSE.NOTES",
		549 : "ACTIVATE.NOTES",
		620 : "PROTECT.REVISIONS",
		621 : "UNPROTECT.REVISIONS",
		647 : "OPTIONS.ME",
		653 : "WEB.PUBLISH",
		667 : "NEWWEBQUERY",
		673 : "PIVOT.TABLE.CHART",
		753 : "OPTIONS.SAVE",
		755 : "OPTIONS.SPELL",
		808 : "HIDEALL.INKANNOTS"
	};
	var cv = {
		0 : "COUNT",
		1 : "IF",
		2 : "ISNA",
		3 : "ISERROR",
		4 : "SUM",
		5 : "AVERAGE",
		6 : "MIN",
		7 : "MAX",
		8 : "ROW",
		9 : "COLUMN",
		10 : "NA",
		11 : "NPV",
		12 : "STDEV",
		13 : "DOLLAR",
		14 : "FIXED",
		15 : "SIN",
		16 : "COS",
		17 : "TAN",
		18 : "ATAN",
		19 : "PI",
		20 : "SQRT",
		21 : "EXP",
		22 : "LN",
		23 : "LOG10",
		24 : "ABS",
		25 : "INT",
		26 : "SIGN",
		27 : "ROUND",
		28 : "LOOKUP",
		29 : "INDEX",
		30 : "REPT",
		31 : "MID",
		32 : "LEN",
		33 : "VALUE",
		34 : "TRUE",
		35 : "FALSE",
		36 : "AND",
		37 : "OR",
		38 : "NOT",
		39 : "MOD",
		40 : "DCOUNT",
		41 : "DSUM",
		42 : "DAVERAGE",
		43 : "DMIN",
		44 : "DMAX",
		45 : "DSTDEV",
		46 : "VAR",
		47 : "DVAR",
		48 : "TEXT",
		49 : "LINEST",
		50 : "TREND",
		51 : "LOGEST",
		52 : "GROWTH",
		53 : "GOTO",
		54 : "HALT",
		55 : "RETURN",
		56 : "PV",
		57 : "FV",
		58 : "NPER",
		59 : "PMT",
		60 : "RATE",
		61 : "MIRR",
		62 : "IRR",
		63 : "RAND",
		64 : "MATCH",
		65 : "DATE",
		66 : "TIME",
		67 : "DAY",
		68 : "MONTH",
		69 : "YEAR",
		70 : "WEEKDAY",
		71 : "HOUR",
		72 : "MINUTE",
		73 : "SECOND",
		74 : "NOW",
		75 : "AREAS",
		76 : "ROWS",
		77 : "COLUMNS",
		78 : "OFFSET",
		79 : "ABSREF",
		80 : "RELREF",
		81 : "ARGUMENT",
		82 : "SEARCH",
		83 : "TRANSPOSE",
		84 : "ERROR",
		85 : "STEP",
		86 : "TYPE",
		87 : "ECHO",
		88 : "SET.NAME",
		89 : "CALLER",
		90 : "DEREF",
		91 : "WINDOWS",
		92 : "SERIES",
		93 : "DOCUMENTS",
		94 : "ACTIVE.CELL",
		95 : "SELECTION",
		96 : "RESULT",
		97 : "ATAN2",
		98 : "ASIN",
		99 : "ACOS",
		100 : "CHOOSE",
		101 : "HLOOKUP",
		102 : "VLOOKUP",
		103 : "LINKS",
		104 : "INPUT",
		105 : "ISREF",
		106 : "GET.FORMULA",
		107 : "GET.NAME",
		108 : "SET.VALUE",
		109 : "LOG",
		110 : "EXEC",
		111 : "CHAR",
		112 : "LOWER",
		113 : "UPPER",
		114 : "PROPER",
		115 : "LEFT",
		116 : "RIGHT",
		117 : "EXACT",
		118 : "TRIM",
		119 : "REPLACE",
		120 : "SUBSTITUTE",
		121 : "CODE",
		122 : "NAMES",
		123 : "DIRECTORY",
		124 : "FIND",
		125 : "CELL",
		126 : "ISERR",
		127 : "ISTEXT",
		128 : "ISNUMBER",
		129 : "ISBLANK",
		130 : "T",
		131 : "N",
		132 : "FOPEN",
		133 : "FCLOSE",
		134 : "FSIZE",
		135 : "FREADLN",
		136 : "FREAD",
		137 : "FWRITELN",
		138 : "FWRITE",
		139 : "FPOS",
		140 : "DATEVALUE",
		141 : "TIMEVALUE",
		142 : "SLN",
		143 : "SYD",
		144 : "DDB",
		145 : "GET.DEF",
		146 : "REFTEXT",
		147 : "TEXTREF",
		148 : "INDIRECT",
		149 : "REGISTER",
		150 : "CALL",
		151 : "ADD.BAR",
		152 : "ADD.MENU",
		153 : "ADD.COMMAND",
		154 : "ENABLE.COMMAND",
		155 : "CHECK.COMMAND",
		156 : "RENAME.COMMAND",
		157 : "SHOW.BAR",
		158 : "DELETE.MENU",
		159 : "DELETE.COMMAND",
		160 : "GET.CHART.ITEM",
		161 : "DIALOG.BOX",
		162 : "CLEAN",
		163 : "MDETERM",
		164 : "MINVERSE",
		165 : "MMULT",
		166 : "FILES",
		167 : "IPMT",
		168 : "PPMT",
		169 : "COUNTA",
		170 : "CANCEL.KEY",
		171 : "FOR",
		172 : "WHILE",
		173 : "BREAK",
		174 : "NEXT",
		175 : "INITIATE",
		176 : "REQUEST",
		177 : "POKE",
		178 : "EXECUTE",
		179 : "TERMINATE",
		180 : "RESTART",
		181 : "HELP",
		182 : "GET.BAR",
		183 : "PRODUCT",
		184 : "FACT",
		185 : "GET.CELL",
		186 : "GET.WORKSPACE",
		187 : "GET.WINDOW",
		188 : "GET.DOCUMENT",
		189 : "DPRODUCT",
		190 : "ISNONTEXT",
		191 : "GET.NOTE",
		192 : "NOTE",
		193 : "STDEVP",
		194 : "VARP",
		195 : "DSTDEVP",
		196 : "DVARP",
		197 : "TRUNC",
		198 : "ISLOGICAL",
		199 : "DCOUNTA",
		200 : "DELETE.BAR",
		201 : "UNREGISTER",
		204 : "USDOLLAR",
		205 : "FINDB",
		206 : "SEARCHB",
		207 : "REPLACEB",
		208 : "LEFTB",
		209 : "RIGHTB",
		210 : "MIDB",
		211 : "LENB",
		212 : "ROUNDUP",
		213 : "ROUNDDOWN",
		214 : "ASC",
		215 : "DBCS",
		216 : "RANK",
		219 : "ADDRESS",
		220 : "DAYS360",
		221 : "TODAY",
		222 : "VDB",
		223 : "ELSE",
		224 : "ELSE.IF",
		225 : "END.IF",
		226 : "FOR.CELL",
		227 : "MEDIAN",
		228 : "SUMPRODUCT",
		229 : "SINH",
		230 : "COSH",
		231 : "TANH",
		232 : "ASINH",
		233 : "ACOSH",
		234 : "ATANH",
		235 : "DGET",
		236 : "CREATE.OBJECT",
		237 : "VOLATILE",
		238 : "LAST.ERROR",
		239 : "CUSTOM.UNDO",
		240 : "CUSTOM.REPEAT",
		241 : "FORMULA.CONVERT",
		242 : "GET.LINK.INFO",
		243 : "TEXT.BOX",
		244 : "INFO",
		245 : "GROUP",
		246 : "GET.OBJECT",
		247 : "DB",
		248 : "PAUSE",
		251 : "RESUME",
		252 : "FREQUENCY",
		253 : "ADD.TOOLBAR",
		254 : "DELETE.TOOLBAR",
		255 : "User",
		256 : "RESET.TOOLBAR",
		257 : "EVALUATE",
		258 : "GET.TOOLBAR",
		259 : "GET.TOOL",
		260 : "SPELLING.CHECK",
		261 : "ERROR.TYPE",
		262 : "APP.TITLE",
		263 : "WINDOW.TITLE",
		264 : "SAVE.TOOLBAR",
		265 : "ENABLE.TOOL",
		266 : "PRESS.TOOL",
		267 : "REGISTER.ID",
		268 : "GET.WORKBOOK",
		269 : "AVEDEV",
		270 : "BETADIST",
		271 : "GAMMALN",
		272 : "BETAINV",
		273 : "BINOMDIST",
		274 : "CHIDIST",
		275 : "CHIINV",
		276 : "COMBIN",
		277 : "CONFIDENCE",
		278 : "CRITBINOM",
		279 : "EVEN",
		280 : "EXPONDIST",
		281 : "FDIST",
		282 : "FINV",
		283 : "FISHER",
		284 : "FISHERINV",
		285 : "FLOOR",
		286 : "GAMMADIST",
		287 : "GAMMAINV",
		288 : "CEILING",
		289 : "HYPGEOMDIST",
		290 : "LOGNORMDIST",
		291 : "LOGINV",
		292 : "NEGBINOMDIST",
		293 : "NORMDIST",
		294 : "NORMSDIST",
		295 : "NORMINV",
		296 : "NORMSINV",
		297 : "STANDARDIZE",
		298 : "ODD",
		299 : "PERMUT",
		300 : "POISSON",
		301 : "TDIST",
		302 : "WEIBULL",
		303 : "SUMXMY2",
		304 : "SUMX2MY2",
		305 : "SUMX2PY2",
		306 : "CHITEST",
		307 : "CORREL",
		308 : "COVAR",
		309 : "FORECAST",
		310 : "FTEST",
		311 : "INTERCEPT",
		312 : "PEARSON",
		313 : "RSQ",
		314 : "STEYX",
		315 : "SLOPE",
		316 : "TTEST",
		317 : "PROB",
		318 : "DEVSQ",
		319 : "GEOMEAN",
		320 : "HARMEAN",
		321 : "SUMSQ",
		322 : "KURT",
		323 : "SKEW",
		324 : "ZTEST",
		325 : "LARGE",
		326 : "SMALL",
		327 : "QUARTILE",
		328 : "PERCENTILE",
		329 : "PERCENTRANK",
		330 : "MODE",
		331 : "TRIMMEAN",
		332 : "TINV",
		334 : "MOVIE.COMMAND",
		335 : "GET.MOVIE",
		336 : "CONCATENATE",
		337 : "POWER",
		338 : "PIVOT.ADD.DATA",
		339 : "GET.PIVOT.TABLE",
		340 : "GET.PIVOT.FIELD",
		341 : "GET.PIVOT.ITEM",
		342 : "RADIANS",
		343 : "DEGREES",
		344 : "SUBTOTAL",
		345 : "SUMIF",
		346 : "COUNTIF",
		347 : "COUNTBLANK",
		348 : "SCENARIO.GET",
		349 : "OPTIONS.LISTS.GET",
		350 : "ISPMT",
		351 : "DATEDIF",
		352 : "DATESTRING",
		353 : "NUMBERSTRING",
		354 : "ROMAN",
		355 : "OPEN.DIALOG",
		356 : "SAVE.DIALOG",
		357 : "VIEW.GET",
		358 : "GETPIVOTDATA",
		359 : "HYPERLINK",
		360 : "PHONETIC",
		361 : "AVERAGEA",
		362 : "MAXA",
		363 : "MINA",
		364 : "STDEVPA",
		365 : "VARPA",
		366 : "STDEVA",
		367 : "VARA",
		368 : "BAHTTEXT",
		369 : "THAIDAYOFWEEK",
		370 : "THAIDIGIT",
		371 : "THAIMONTHOFYEAR",
		372 : "THAINUMSOUND",
		373 : "THAINUMSTRING",
		374 : "THAISTRINGLENGTH",
		375 : "ISTHAIDIGIT",
		376 : "ROUNDBAHTDOWN",
		377 : "ROUNDBAHTUP",
		378 : "THAIYEAR",
		379 : "RTD",
		380 : "CUBEVALUE",
		381 : "CUBEMEMBER",
		382 : "CUBEMEMBERPROPERTY",
		383 : "CUBERANKEDMEMBER",
		384 : "HEX2BIN",
		385 : "HEX2DEC",
		386 : "HEX2OCT",
		387 : "DEC2BIN",
		388 : "DEC2HEX",
		389 : "DEC2OCT",
		390 : "OCT2BIN",
		391 : "OCT2HEX",
		392 : "OCT2DEC",
		393 : "BIN2DEC",
		394 : "BIN2OCT",
		395 : "BIN2HEX",
		396 : "IMSUB",
		397 : "IMDIV",
		398 : "IMPOWER",
		399 : "IMABS",
		400 : "IMSQRT",
		401 : "IMLN",
		402 : "IMLOG2",
		403 : "IMLOG10",
		404 : "IMSIN",
		405 : "IMCOS",
		406 : "IMEXP",
		407 : "IMARGUMENT",
		408 : "IMCONJUGATE",
		409 : "IMAGINARY",
		410 : "IMREAL",
		411 : "COMPLEX",
		412 : "IMSUM",
		413 : "IMPRODUCT",
		414 : "SERIESSUM",
		415 : "FACTDOUBLE",
		416 : "SQRTPI",
		417 : "QUOTIENT",
		418 : "DELTA",
		419 : "GESTEP",
		420 : "ISEVEN",
		421 : "ISODD",
		422 : "MROUND",
		423 : "ERF",
		424 : "ERFC",
		425 : "BESSELJ",
		426 : "BESSELK",
		427 : "BESSELY",
		428 : "BESSELI",
		429 : "XIRR",
		430 : "XNPV",
		431 : "PRICEMAT",
		432 : "YIELDMAT",
		433 : "INTRATE",
		434 : "RECEIVED",
		435 : "DISC",
		436 : "PRICEDISC",
		437 : "YIELDDISC",
		438 : "TBILLEQ",
		439 : "TBILLPRICE",
		440 : "TBILLYIELD",
		441 : "PRICE",
		442 : "YIELD",
		443 : "DOLLARDE",
		444 : "DOLLARFR",
		445 : "NOMINAL",
		446 : "EFFECT",
		447 : "CUMPRINC",
		448 : "CUMIPMT",
		449 : "EDATE",
		450 : "EOMONTH",
		451 : "YEARFRAC",
		452 : "COUPDAYBS",
		453 : "COUPDAYS",
		454 : "COUPDAYSNC",
		455 : "COUPNCD",
		456 : "COUPNUM",
		457 : "COUPPCD",
		458 : "DURATION",
		459 : "MDURATION",
		460 : "ODDLPRICE",
		461 : "ODDLYIELD",
		462 : "ODDFPRICE",
		463 : "ODDFYIELD",
		464 : "RANDBETWEEN",
		465 : "WEEKNUM",
		466 : "AMORDEGRC",
		467 : "AMORLINC",
		468 : "CONVERT",
		724 : "SHEETJS",
		469 : "ACCRINT",
		470 : "ACCRINTM",
		471 : "WORKDAY",
		472 : "NETWORKDAYS",
		473 : "GCD",
		474 : "MULTINOMIAL",
		475 : "LCM",
		476 : "FVSCHEDULE",
		477 : "CUBEKPIMEMBER",
		478 : "CUBESET",
		479 : "CUBESETCOUNT",
		480 : "IFERROR",
		481 : "COUNTIFS",
		482 : "SUMIFS",
		483 : "AVERAGEIF",
		484 : "AVERAGEIFS"
	};
	var uv = {
		2 : 1,
		3 : 1,
		10 : 0,
		15 : 1,
		16 : 1,
		17 : 1,
		18 : 1,
		19 : 0,
		20 : 1,
		21 : 1,
		22 : 1,
		23 : 1,
		24 : 1,
		25 : 1,
		26 : 1,
		27 : 2,
		30 : 2,
		31 : 3,
		32 : 1,
		33 : 1,
		34 : 0,
		35 : 0,
		38 : 1,
		39 : 2,
		40 : 3,
		41 : 3,
		42 : 3,
		43 : 3,
		44 : 3,
		45 : 3,
		47 : 3,
		48 : 2,
		53 : 1,
		61 : 3,
		63 : 0,
		65 : 3,
		66 : 3,
		67 : 1,
		68 : 1,
		69 : 1,
		70 : 1,
		71 : 1,
		72 : 1,
		73 : 1,
		74 : 0,
		75 : 1,
		76 : 1,
		77 : 1,
		79 : 2,
		80 : 2,
		83 : 1,
		85 : 0,
		86 : 1,
		89 : 0,
		90 : 1,
		94 : 0,
		95 : 0,
		97 : 2,
		98 : 1,
		99 : 1,
		101 : 3,
		102 : 3,
		105 : 1,
		106 : 1,
		108 : 2,
		111 : 1,
		112 : 1,
		113 : 1,
		114 : 1,
		117 : 2,
		118 : 1,
		119 : 4,
		121 : 1,
		126 : 1,
		127 : 1,
		128 : 1,
		129 : 1,
		130 : 1,
		131 : 1,
		133 : 1,
		134 : 1,
		135 : 1,
		136 : 2,
		137 : 2,
		138 : 2,
		140 : 1,
		141 : 1,
		142 : 3,
		143 : 4,
		144 : 4,
		161 : 1,
		162 : 1,
		163 : 1,
		164 : 1,
		165 : 2,
		172 : 1,
		175 : 2,
		176 : 2,
		177 : 3,
		178 : 2,
		179 : 1,
		184 : 1,
		186 : 1,
		189 : 3,
		190 : 1,
		195 : 3,
		196 : 3,
		197 : 1,
		198 : 1,
		199 : 3,
		201 : 1,
		207 : 4,
		210 : 3,
		211 : 1,
		212 : 2,
		213 : 2,
		214 : 1,
		215 : 1,
		225 : 0,
		229 : 1,
		230 : 1,
		231 : 1,
		232 : 1,
		233 : 1,
		234 : 1,
		235 : 3,
		244 : 1,
		247 : 4,
		252 : 2,
		257 : 1,
		261 : 1,
		271 : 1,
		273 : 4,
		274 : 2,
		275 : 2,
		276 : 2,
		277 : 3,
		278 : 3,
		279 : 1,
		280 : 3,
		281 : 3,
		282 : 3,
		283 : 1,
		284 : 1,
		285 : 2,
		286 : 4,
		287 : 3,
		288 : 2,
		289 : 4,
		290 : 3,
		291 : 3,
		292 : 3,
		293 : 4,
		294 : 1,
		295 : 3,
		296 : 1,
		297 : 3,
		298 : 1,
		299 : 2,
		300 : 3,
		301 : 3,
		302 : 4,
		303 : 2,
		304 : 2,
		305 : 2,
		306 : 2,
		307 : 2,
		308 : 2,
		309 : 3,
		310 : 2,
		311 : 2,
		312 : 2,
		313 : 2,
		314 : 2,
		315 : 2,
		316 : 4,
		325 : 2,
		326 : 2,
		327 : 2,
		328 : 2,
		331 : 2,
		332 : 2,
		337 : 2,
		342 : 1,
		343 : 1,
		346 : 2,
		347 : 1,
		350 : 4,
		351 : 3,
		352 : 1,
		353 : 2,
		360 : 1,
		368 : 1,
		369 : 1,
		370 : 1,
		371 : 1,
		372 : 1,
		373 : 1,
		374 : 1,
		375 : 1,
		376 : 1,
		377 : 1,
		378 : 1,
		382 : 3,
		385 : 1,
		392 : 1,
		393 : 1,
		396 : 2,
		397 : 2,
		398 : 2,
		399 : 1,
		400 : 1,
		401 : 1,
		402 : 1,
		403 : 1,
		404 : 1,
		405 : 1,
		406 : 1,
		407 : 1,
		408 : 1,
		409 : 1,
		410 : 1,
		414 : 4,
		415 : 1,
		416 : 1,
		417 : 2,
		420 : 1,
		421 : 1,
		422 : 2,
		424 : 1,
		425 : 2,
		426 : 2,
		427 : 2,
		428 : 2,
		430 : 3,
		438 : 3,
		439 : 3,
		440 : 3,
		443 : 2,
		444 : 2,
		445 : 2,
		446 : 2,
		447 : 6,
		448 : 6,
		449 : 2,
		450 : 2,
		464 : 2,
		468 : 3,
		476 : 2,
		479 : 1,
		480 : 2,
		65535 : 0
	};
	function hv(e) {
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
	function dv(e) {
		var r = "of:=" + e.replace(th, "$1[.$2$3$4$5]").replace(/\]:\[/g, ":");
		return r.replace(/;/g, "|").replace(/,/g, ";")
	}
	function vv(e) {
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
		return [t, r[0].split(".")[1] + (r.length > 1 ? ":" + (r[1].split(".")[1] || r[1].split(".")[0]) : "")]
	}
	function pv(e) {
		return e.replace(/!/, ".").replace(/:/, ":.")
	}
	var mv = {};
	var bv = {};
	var gv = typeof Map !== "undefined";
	function wv(e, r, t) {
		var a = 0,
		n = e.length;
		if (t) {
			if (gv ? t.has(r) : Object.prototype.hasOwnProperty.call(t, r)) {
				var i = gv ? t.get(r) : t[r];
				for (; a < i.length; ++a) {
					if (e[i[a]].t === r) {
						e.Count++;
						return i[a]
					}
				}
			}
		} else for (; a < n; ++a) {
			if (e[a].t === r) {
				e.Count++;
				return a
			}
		}
		e[n] = {
			t: r
		};
		e.Count++;
		e.Unique++;
		if (t) {
			if (gv) {
				if (!t.has(r)) t.set(r, []);
				t.get(r).push(n)
			} else {
				if (!Object.prototype.hasOwnProperty.call(t, r)) t[r] = [];
				t[r].push(n)
			}
		}
		return n
	}
	function kv(e, r) {
		var t = {
			min: e + 1,
			max: e + 1
		};
		var a = -1;
		if (r.MDW) qo = r.MDW;
		if (r.width != null) t.customWidth = 1;
		else if (r.wpx != null) a = ec(r.wpx);
		else if (r.wch != null) a = r.wch;
		if (a > -1) {
			t.width = rc(a);
			t.customWidth = 1
		} else if (r.width != null) t.width = r.width;
		if (r.hidden) t.hidden = true;
		if (r.level != null) {
			t.outlineLevel = t.level = r.level
		}
		return t
	}
	function Tv(e, r) {
		if (!e) return;
		var t = [.7, .7, .75, .75, .3, .3];
		if (r == "xlml") t = [1, 1, 1, 1, .5, .5];
		if (e.left == null) e.left = t[0];
		if (e.right == null) e.right = t[1];
		if (e.top == null) e.top = t[2];
		if (e.bottom == null) e.bottom = t[3];
		if (e.header == null) e.header = t[4];
		if (e.footer == null) e.footer = t[5]
	}
	function yv(e, r, t) {
		var a = t.revssf[r.z != null ? r.z: "General"];
		var n = 60,
		i = e.length;
		if (a == null && t.ssf) {
			for (; n < 392; ++n) if (t.ssf[n] == null) {
				Je(r.z, n);
				t.ssf[n] = r.z;
				t.revssf[r.z] = a = n;
				break
			}
		}
		for (n = 0; n != i; ++n) if (e[n].numFmtId === a) return n;
		e[i] = {
			numFmtId: a,
			fontId: 0,
			fillId: 0,
			borderId: 0,
			xfId: 0,
			applyNumberFormat: 1
		};
		return i
	}
	function Ev(e, r, t, a, n, i, s) {
		try {
			if (a.cellNF) e.z = q[r]
		} catch(f) {
			if (a.WTF) throw f
		}
		if (e.t === "z" && !a.cellStyles) return;
		if (e.t === "d" && typeof e.v === "string") e.v = wr(e.v);
		if ((!a || a.cellText !== false) && e.t !== "z") try {
			if (q[r] == null) Je(Ge[r] || "General", r);
			if (e.t === "e") e.w = e.w || ti[e.v];
			else if (r === 0) {
				if (e.t === "n") {
					if ((e.v | 0) === e.v) e.w = e.v.toString(10);
					else e.w = le(e.v)
				} else if (e.t === "d") {
					var l = dr(e.v, !!s);
					if ((l | 0) === l) e.w = l.toString(10);
					else e.w = le(l)
				} else if (e.v === undefined) return "";
				else e.w = oe(e.v, bv)
			} else if (e.t === "d") e.w = ze(r, dr(e.v, !!s), bv);
			else e.w = ze(r, e.v, bv)
		} catch(f) {
			if (a.WTF) throw f
		}
		if (!a.cellStyles) return;
		if (t != null) try {
			e.s = i.Fills[t];
			if (e.s.fgColor && e.s.fgColor.theme && !e.s.fgColor.rgb) {
				e.s.fgColor.rgb = Ko(n.themeElements.clrScheme[e.s.fgColor.theme].rgb, e.s.fgColor.tint || 0);
				if (a.WTF) e.s.fgColor.raw_rgb = n.themeElements.clrScheme[e.s.fgColor.theme].rgb
			}
			if (e.s.bgColor && e.s.bgColor.theme) {
				e.s.bgColor.rgb = Ko(n.themeElements.clrScheme[e.s.bgColor.theme].rgb, e.s.bgColor.tint || 0);
				if (a.WTF) e.s.bgColor.raw_rgb = n.themeElements.clrScheme[e.s.bgColor.theme].rgb
			}
		} catch(f) {
			if (a.WTF && i.Fills) throw f
		}
	}
	function _v(e, r, t) {
		if (e && e["!ref"]) {
			var a = Ga(e["!ref"]);
			if (a.e.c < a.s.c || a.e.r < a.s.r) throw new Error("Bad range (" + t + "): " + e["!ref"])
		}
	}
	function Sv(e, r) {
		var t = Ga(r);
		if (t.s.r <= t.e.r && t.s.c <= t.e.c && t.s.r >= 0 && t.s.c >= 0) e["!ref"] = Va(t)
	}
	var xv = /<(?:\w:)?mergeCell ref=["'][A-Z0-9:]+['"]\s*[\/]?>/g;
	var Av = /<(?:\w+:)?sheetData[^>]*>([\s\S]*)<\/(?:\w+:)?sheetData>/;
	var Cv = /<(?:\w:)?hyperlink [^>]*>/gm;
	var Rv = /"(\w*:\w*)"/;
	var Ov = /<(?:\w:)?col\b[^>]*[\/]?>/g;
	var Iv = /<(?:\w:)?autoFilter[^>]*([\/]|>([\s\S]*)<\/(?:\w:)?autoFilter)>/g;
	var Nv = /<(?:\w:)?pageMargins[^>]*\/>/g;
	var Fv = /<(?:\w:)?sheetPr\b(?:[^>a-z][^>]*)?\/>/;
	var Dv = /<(?:\w:)?sheetPr[^>]*(?:[\/]|>([\s\S]*)<\/(?:\w:)?sheetPr)>/;
	var Pv = /<(?:\w:)?sheetViews[^>]*(?:[\/]|>([\s\S]*)<\/(?:\w:)?sheetViews)>/;
	function Lv(e, r, t, a, n, i, s) {
		if (!e) return e;
		if (!a) a = {
			"!id": {}
		};
		if (g != null && r.dense == null) r.dense = g;
		var f = {};
		if (r.dense) f["!data"] = [];
		var l = {
			s: {
				r: 2e6,
				c: 2e6
			},
			e: {
				r: 0,
				c: 0
			}
		};
		var o = "",
		c = "";
		var u = e.match(Av);
		if (u) {
			o = e.slice(0, u.index);
			c = e.slice(u.index + u[0].length)
		} else o = c = e;
		var h = o.match(Fv);
		if (h) Uv(h[0], f, n, t);
		else if (h = o.match(Dv)) Bv(h[0], h[1] || "", f, n, t, s, i);
		var d = (o.match(/<(?:\w*:)?dimension/) || {
			index: -1
		}).index;
		if (d > 0) {
			var v = o.slice(d, d + 50).match(Rv);
			if (v && !(r && r.nodim)) Sv(f, v[1])
		}
		var p = o.match(Pv);
		if (p && p[1]) qv(p[1], n);
		var m = [];
		if (r.cellStyles) {
			var b = o.match(Ov);
			if (b) jv(m, b)
		}
		if (u) rp(u[1], f, r, l, i, s, n);
		var w = c.match(Iv);
		if (w) f["!autofilter"] = Yv(w[0]);
		var k = [];
		var T = c.match(xv);
		if (T) for (d = 0; d != T.length; ++d) k[d] = Ga(T[d].slice(T[d].indexOf('"') + 1));
		var y = c.match(Cv);
		if (y) $v(f, y, a);
		var E = c.match(Nv);
		if (E) f["!margins"] = Xv(rt(E[0]));
		var _;
		if (_ = c.match(/legacyDrawing r:id="(.*?)"/)) f["!legrel"] = _[1];
		if (r && r.nodim) l.s.c = l.s.r = 0;
		if (!f["!ref"] && l.e.c >= l.s.c && l.e.r >= l.s.r) f["!ref"] = Va(l);
		if (r.sheetRows > 0 && f["!ref"]) {
			var S = Ga(f["!ref"]);
			if (r.sheetRows <= +S.e.r) {
				S.e.r = r.sheetRows - 1;
				if (S.e.r > l.e.r) S.e.r = l.e.r;
				if (S.e.r < S.s.r) S.s.r = S.e.r;
				if (S.e.c > l.e.c) S.e.c = l.e.c;
				if (S.e.c < S.s.c) S.s.c = S.e.c;
				f["!fullref"] = f["!ref"];
				f["!ref"] = Va(S)
			}
		}
		if (m.length > 0) f["!cols"] = m;
		if (k.length > 0) f["!merges"] = k;
		if (a["!id"][f["!legrel"]]) f["!legdrawel"] = a["!id"][f["!legrel"]];
		return f
	}
	function Mv(e) {
		if (e.length === 0) return "";
		var r = '<mergeCells count="' + e.length + '">';
		for (var t = 0; t != e.length; ++t) r += '<mergeCell ref="' + Va(e[t]) + '"/>';
		return r + "</mergeCells>"
	}
	function Uv(e, r, t, a) {
		var n = rt(e);
		if (!t.Sheets[a]) t.Sheets[a] = {};
		if (n.codeName) t.Sheets[a].CodeName = it(kt(n.codeName))
	}
	function Bv(e, r, t, a, n) {
		Uv(e.slice(0, e.indexOf(">")), t, a, n)
	}
	function Wv(e, r, t, a, n) {
		var i = false;
		var s = {},
		f = null;
		if (a.bookType !== "xlsx" && r.vbaraw) {
			var l = r.SheetNames[t];
			try {
				if (r.Workbook) l = r.Workbook.Sheets[t].CodeName || l
			} catch(o) {}
			i = true;
			s.codeName = Tt(lt(l))
		}
		if (e && e["!outline"]) {
			var c = {
				summaryBelow: 1,
				summaryRight: 1
			};
			if (e["!outline"].above) c.summaryBelow = 0;
			if (e["!outline"].left) c.summaryRight = 0;
			f = (f || "") + It("outlinePr", null, c)
		}
		if (!i && !f) return;
		n[n.length] = It("sheetPr", f, s)
	}
	var zv = ["objects", "scenarios", "selectLockedCells", "selectUnlockedCells"];
	var Hv = ["formatColumns", "formatRows", "formatCells", "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows", "sort", "autoFilter", "pivotTables"];
	function Vv(e) {
		var r = {
			sheet: 1
		};
		zv.forEach(function(t) {
			if (e[t] != null && e[t]) r[t] = "1"
		});
		Hv.forEach(function(t) {
			if (e[t] != null && !e[t]) r[t] = "0"
		});
		if (e.password) r.password = Fo(e.password).toString(16).toUpperCase();
		return It("sheetProtection", null, r)
	}
	function $v(e, r, t) {
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
					TargetMode: "Internal"
				}
			}
			i.Rel = s;
			if (i.tooltip) {
				i.Tooltip = i.tooltip;
				delete i.tooltip
			}
			var f = Ga(i.ref);
			for (var l = f.s.r; l <= f.e.r; ++l) for (var o = f.s.c; o <= f.e.c; ++o) {
				var c = La(o) + Na(l);
				if (a) {
					if (!e["!data"][l]) e["!data"][l] = [];
					if (!e["!data"][l][o]) e["!data"][l][o] = {
						t: "z",
						v: undefined
					};
					e["!data"][l][o].l = i
				} else {
					if (!e[c]) e[c] = {
						t: "z",
						v: undefined
					};
					e[c].l = i
				}
			}
		}
	}
	function Xv(e) {
		var r = {}; ["left", "right", "top", "bottom", "header", "footer"].forEach(function(t) {
			if (e[t]) r[t] = parseFloat(e[t])
		});
		return r
	}
	function Gv(e) {
		Tv(e);
		return It("pageMargins", null, e)
	}
	function jv(e, r) {
		var t = false;
		for (var a = 0; a != r.length; ++a) {
			var n = rt(r[a], true);
			if (n.hidden) n.hidden = pt(n.hidden);
			var i = parseInt(n.min, 10) - 1,
			s = parseInt(n.max, 10) - 1;
			if (n.outlineLevel) n.level = +n.outlineLevel || 0;
			delete n.min;
			delete n.max;
			n.width = +n.width;
			if (!t && n.width) {
				t = true;
				ac(n.width)
			}
			nc(n);
			while (i <= s) e[i++] = Tr(n)
		}
	}
	function Kv(e, r) {
		var t = ["<cols>"],
		a;
		for (var n = 0; n != r.length; ++n) {
			if (! (a = r[n])) continue;
			t[t.length] = It("col", null, kv(n, a))
		}
		t[t.length] = "</cols>";
		return t.join("")
	}
	function Yv(e) {
		var r = {
			ref: (e.match(/ref="([^"]*)"/) || [])[1]
		};
		return r
	}
	function Zv(e, r, t, a) {
		var n = typeof e.ref == "string" ? e.ref: Va(e.ref);
		if (!t.Workbook) t.Workbook = {
			Sheets: []
		};
		if (!t.Workbook.Names) t.Workbook.Names = [];
		var i = t.Workbook.Names;
		var s = Ha(n);
		if (s.s.r == s.e.r) {
			s.e.r = Ha(r["!ref"]).e.r;
			n = Va(s)
		}
		for (var f = 0; f < i.length; ++f) {
			var l = i[f];
			if (l.Name != "_xlnm._FilterDatabase") continue;
			if (l.Sheet != a) continue;
			l.Ref = Xa(t.SheetNames[a]) + "!" + $a(n);
			break
		}
		if (f == i.length) i.push({
			Name: "_xlnm._FilterDatabase",
			Sheet: a,
			Ref: "'" + t.SheetNames[a] + "'!" + n
		});
		return It("autoFilter", null, {
			ref: n
		})
	}
	var Jv = /<(?:\w:)?sheetView(?:[^>a-z][^>]*)?\/?>/g;
	function qv(e, r) {
		if (!r.Views) r.Views = [{}]; (e.match(Jv) || []).forEach(function(e, t) {
			var a = rt(e);
			if (!r.Views[t]) r.Views[t] = {};
			if ( + a.zoomScale) r.Views[t].zoom = +a.zoomScale;
			if (a.rightToLeft && pt(a.rightToLeft)) r.Views[t].RTL = true
		})
	}
	function Qv(e, r, t, a) {
		var n = {
			workbookViewId: "0"
		};
		if ((((a || {}).Workbook || {}).Views || [])[0]) n.rightToLeft = a.Workbook.Views[0].RTL ? "1": "0";
		return It("sheetViews", It("sheetView", null, n), {})
	}
	function ep(e, r, t, a, n, i, s) {
		if (e.c) t["!comments"].push([r, e.c]);
		if ((e.v === undefined || e.t === "z" && !(a || {}).sheetStubs) && typeof e.f !== "string" && typeof e.z == "undefined") return "";
		var f = "";
		var l = e.t,
		o = e.v;
		if (e.t !== "z") switch (e.t) {
		case "b":
			f = e.v ? "1": "0";
			break;
		case "n":
			if (isNaN(e.v)) {
				e.t = "e";
				f = ti[e.v = 36]
			} else if (!isFinite(e.v)) {
				e.t = "e";
				f = ti[e.v = 7]
			} else f = "" + e.v;
			break;
		case "e":
			f = ti[e.v];
			break;
		case "d":
			if (a && a.cellDates) {
				var c = wr(e.v, s);
				f = c.toISOString();
				if (c.getUTCFullYear() < 1900) f = f.slice(f.indexOf("T") + 1).replace("Z", "")
			} else {
				e = Tr(e);
				e.t = "n";
				f = "" + (e.v = dr(wr(e.v, s), s))
			}
			if (typeof e.z === "undefined") e.z = q[14];
			break;
		default:
			f = e.v;
			break;
		}
		var u = e.t == "z" || e.v == null ? "": Rt("v", lt(f)),
		h = {
			r: r
		};
		var d = yv(a.cellXfs, e, a);
		if (d !== 0) h.s = d;
		switch (e.t) {
		case "n":
			break;
		case "d":
			h.t = "d";
			break;
		case "b":
			h.t = "b";
			break;
		case "e":
			h.t = "e";
			break;
		case "z":
			break;
		default:
			if (e.v == null) {
				delete e.t;
				break
			}
			if (e.v.length > 32767) throw new Error("Text length must not exceed 32767 characters");
			if (a && a.bookSST) {
				u = Rt("v", "" + wv(a.Strings, e.v, a.revStrings));
				h.t = "s";
				break
			} else h.t = "str";
			break;
		}
		if (e.t != l) {
			e.t = l;
			e.v = o
		}
		if (typeof e.f == "string" && e.f) {
			var v = e.F && e.F.slice(0, r.length) == r ? {
				t: "array",
				ref: e.F
			}: null;
			u = It("f", lt(e.f), v) + (e.v != null ? u: "")
		}
		if (e.l) {
			e.l.display = lt(f);
			t["!links"].push([r, e.l])
		}
		if (e.D) h.cm = 1;
		return It("c", u, h)
	}
	var rp = function() {
		var e = /<(?:\w+:)?c[ \/>]/,
		r = /<\/(?:\w+:)?row>/;
		var t = /r=["']([^"']*)["']/,
		a = /<(?:\w+:)?is>([\S\s]*?)<\/(?:\w+:)?is>/;
		var n = /ref=["']([^"']*)["']/;
		var i = yt("v"),
		s = yt("f");
		return function f(l, o, c, u, h, d, v) {
			var p = 0,
			m = "",
			b = [],
			g = [],
			w = 0,
			k = 0,
			T = 0,
			y = "",
			E;
			var _, S = 0,
			x = 0;
			var A, C;
			var R = 0,
			O = 0;
			var I = Array.isArray(d.CellXf),
			N;
			var F = [];
			var D = [];
			var P = o["!data"] != null;
			var L = [],
			M = {},
			U = false;
			var B = !!c.sheetStubs;
			var W = !!((v || {}).WBProps || {}).date1904;
			for (var z = l.split(r), H = 0, V = z.length; H != V; ++H) {
				m = z[H].trim();
				var $ = m.length;
				if ($ === 0) continue;
				var X = 0;
				e: for (p = 0; p < $; ++p) switch (m[p]) {
				case ">":
					if (m[p - 1] != "/") {++p;
						break e
					}
					if (c && c.cellStyles) {
						_ = rt(m.slice(X, p), true);
						S = _.r != null ? parseInt(_.r, 10) : S + 1;
						x = -1;
						if (c.sheetRows && c.sheetRows < S) continue;
						M = {};
						U = false;
						if (_.ht) {
							U = true;
							M.hpt = parseFloat(_.ht);
							M.hpx = lc(M.hpt)
						}
						if (_.hidden && pt(_.hidden)) {
							U = true;
							M.hidden = true
						}
						if (_.outlineLevel != null) {
							U = true;
							M.level = +_.outlineLevel
						}
						if (U) L[S - 1] = M
					}
					break;
				case "<":
					X = p;
					break;
				}
				if (X >= p) break;
				_ = rt(m.slice(X, p), true);
				S = _.r != null ? parseInt(_.r, 10) : S + 1;
				x = -1;
				if (c.sheetRows && c.sheetRows < S) continue;
				if (!c.nodim) {
					if (u.s.r > S - 1) u.s.r = S - 1;
					if (u.e.r < S - 1) u.e.r = S - 1
				}
				if (c && c.cellStyles) {
					M = {};
					U = false;
					if (_.ht) {
						U = true;
						M.hpt = parseFloat(_.ht);
						M.hpx = lc(M.hpt)
					}
					if (_.hidden && pt(_.hidden)) {
						U = true;
						M.hidden = true
					}
					if (_.outlineLevel != null) {
						U = true;
						M.level = +_.outlineLevel
					}
					if (U) L[S - 1] = M
				}
				b = m.slice(p).split(e);
				for (var G = 0; G != b.length; ++G) if (b[G].trim().charAt(0) != "<") break;
				b = b.slice(G);
				for (p = 0; p != b.length; ++p) {
					m = b[p].trim();
					if (m.length === 0) continue;
					g = m.match(t);
					w = p;
					k = 0;
					T = 0;
					m = "<c " + (m.slice(0, 1) == "<" ? ">": "") + m;
					if (g != null && g.length === 2) {
						w = 0;
						y = g[1];
						for (k = 0; k != y.length; ++k) {
							if ((T = y.charCodeAt(k) - 64) < 1 || T > 26) break;
							w = 26 * w + T
						}--w;
						x = w
					} else++x;
					for (k = 0; k != m.length; ++k) if (m.charCodeAt(k) === 62) break; ++k;
					_ = rt(m.slice(0, k), true);
					if (!_.r) _.r = za({
						r: S - 1,
						c: x
					});
					y = m.slice(k);
					E = {
						t: ""
					};
					if ((g = y.match(i)) != null && g[1] !== "") E.v = it(g[1]);
					if (c.cellFormula) {
						if ((g = y.match(s)) != null) {
							if (g[1] == "") {
								if (g[0].indexOf('t="shared"') > -1) {
									C = rt(g[0]);
									if (D[C.si]) E.f = sh(D[C.si][1], D[C.si][2], _.r)
								}
							} else {
								E.f = it(kt(g[1]), true);
								if (!c.xlfn) E.f = lh(E.f);
								if (g[0].indexOf('t="array"') > -1) {
									E.F = (y.match(n) || [])[1];
									if (E.F.indexOf(":") > -1) F.push([Ga(E.F), E.F])
								} else if (g[0].indexOf('t="shared"') > -1) {
									C = rt(g[0]);
									var j = it(kt(g[1]));
									if (!c.xlfn) j = lh(j);
									D[parseInt(C.si, 10)] = [C, j, _.r]
								}
							}
						} else if (g = y.match(/<f[^>]*\/>/)) {
							C = rt(g[0]);
							if (D[C.si]) E.f = sh(D[C.si][1], D[C.si][2], _.r)
						}
						var K = Wa(_.r);
						for (k = 0; k < F.length; ++k) if (K.r >= F[k][0].s.r && K.r <= F[k][0].e.r) if (K.c >= F[k][0].s.c && K.c <= F[k][0].e.c) E.F = F[k][1]
					}
					if (_.t == null && E.v === undefined) {
						if (E.f || E.F) {
							E.v = 0;
							E.t = "n"
						} else if (!B) continue;
						else E.t = "z"
					} else E.t = _.t || "n";
					if (u.s.c > x) u.s.c = x;
					if (u.e.c < x) u.e.c = x;
					switch (E.t) {
					case "n":
						if (E.v == "" || E.v == null) {
							if (!B) continue;
							E.t = "z"
						} else E.v = parseFloat(E.v);
						break;
					case "s":
						if (typeof E.v == "undefined") {
							if (!B) continue;
							E.t = "z"
						} else {
							A = mv[parseInt(E.v, 10)];
							E.v = A.t;
							E.r = A.r;
							if (c.cellHTML) E.h = A.h
						}
						break;
					case "str":
						E.t = "s";
						E.v = E.v != null ? it(kt(E.v), true) : "";
						if (c.cellHTML) E.h = ut(E.v);
						break;
					case "inlineStr":
						g = y.match(a);
						E.t = "s";
						if (g != null && (A = no(g[1]))) {
							E.v = A.t;
							if (c.cellHTML) E.h = A.h
						} else E.v = "";
						break;
					case "b":
						E.v = pt(E.v);
						break;
					case "d":
						if (c.cellDates) E.v = wr(E.v, W);
						else {
							E.v = dr(wr(E.v, W), W);
							E.t = "n"
						}
						break;
					case "e":
						if (!c || c.cellText !== false) E.w = E.v;
						E.v = ai[E.v];
						break;
					}
					R = O = 0;
					N = null;
					if (I && _.s !== undefined) {
						N = d.CellXf[_.s];
						if (N != null) {
							if (N.numFmtId != null) R = N.numFmtId;
							if (c.cellStyles) {
								if (N.fillId != null) O = N.fillId
							}
						}
					}
					Ev(E, R, O, c, h, d, W);
					if (c.cellDates && I && E.t == "n" && Le(q[R])) {
						E.v = vr(E.v + (W ? 1462 : 0));
						E.t = typeof E.v == "number" ? "n": "d"
					}
					if (_.cm && c.xlmeta) {
						var Y = (c.xlmeta.Cell || [])[ + _.cm - 1];
						if (Y && Y.type == "XLDAPR") E.D = true
					}
					var Z;
					if (c.nodim) {
						Z = Wa(_.r);
						if (u.s.r > Z.r) u.s.r = Z.r;
						if (u.e.r < Z.r) u.e.r = Z.r
					}
					if (P) {
						Z = Wa(_.r);
						if (!o["!data"][Z.r]) o["!data"][Z.r] = [];
						o["!data"][Z.r][Z.c] = E
					} else o[_.r] = E
				}
			}
			if (L.length > 0) o["!rows"] = L
		}
	} ();
	function tp(e, r, t, a) {
		var n = [],
		i = [],
		s = Ga(e["!ref"]),
		f = "",
		l,
		o = "",
		c = [],
		u = 0,
		h = 0,
		d = e["!rows"];
		var v = e["!data"] != null;
		var p = {
			r: o
		},
		m,
		b = -1;
		var g = (((a || {}).Workbook || {}).WBProps || {}).date1904;
		for (h = s.s.c; h <= s.e.c; ++h) c[h] = La(h);
		for (u = s.s.r; u <= s.e.r; ++u) {
			i = [];
			o = Na(u);
			for (h = s.s.c; h <= s.e.c; ++h) {
				l = c[h] + o;
				var w = v ? (e["!data"][u] || [])[h] : e[l];
				if (w === undefined) continue;
				if ((f = ep(w, l, e, r, t, a, g)) != null) i.push(f)
			}
			if (i.length > 0 || d && d[u]) {
				p = {
					r: o
				};
				if (d && d[u]) {
					m = d[u];
					if (m.hidden) p.hidden = 1;
					b = -1;
					if (m.hpx) b = fc(m.hpx);
					else if (m.hpt) b = m.hpt;
					if (b > -1) {
						p.ht = b;
						p.customHeight = 1
					}
					if (m.level) {
						p.outlineLevel = m.level
					}
				}
				n[n.length] = It("row", i.join(""), p)
			}
		}
		if (d) for (; u < d.length; ++u) {
			if (d && d[u]) {
				p = {
					r: u + 1
				};
				m = d[u];
				if (m.hidden) p.hidden = 1;
				b = -1;
				if (m.hpx) b = fc(m.hpx);
				else if (m.hpt) b = m.hpt;
				if (b > -1) {
					p.ht = b;
					p.customHeight = 1
				}
				if (m.level) {
					p.outlineLevel = m.level
				}
				n[n.length] = It("row", "", p)
			}
		}
		return n.join("")
	}
	function ap(e, r, t, a) {
		var n = [Kr, It("worksheet", null, {
			xmlns: Mt[0],
			"xmlns:r": Lt.r
		})];
		var i = t.SheetNames[e],
		s = 0,
		f = "";
		var l = t.Sheets[i];
		if (l == null) l = {};
		var o = l["!ref"] || "A1";
		var c = Ga(o);
		if (c.e.c > 16383 || c.e.r > 1048575) {
			if (r.WTF) throw new Error("Range " + o + " exceeds format limit A1:XFD1048576");
			c.e.c = Math.min(c.e.c, 16383);
			c.e.r = Math.min(c.e.c, 1048575);
			o = Va(c)
		}
		if (!a) a = {};
		l["!comments"] = [];
		var u = [];
		Wv(l, t, e, r, n);
		n[n.length] = It("dimension", null, {
			ref: o
		});
		n[n.length] = Qv(l, r, e, t);
		if (r.sheetFormat) n[n.length] = It("sheetFormatPr", null, {
			defaultRowHeight: r.sheetFormat.defaultRowHeight || "16",
			baseColWidth: r.sheetFormat.baseColWidth || "10",
			outlineLevelRow: r.sheetFormat.outlineLevelRow || "7"
		});
		if (l["!cols"] != null && l["!cols"].length > 0) n[n.length] = Kv(l, l["!cols"]);
		n[s = n.length] = "<sheetData/>";
		l["!links"] = [];
		if (l["!ref"] != null) {
			f = tp(l, r, e, t, a);
			if (f.length > 0) n[n.length] = f
		}
		if (n.length > s + 1) {
			n[n.length] = "</sheetData>";
			n[s] = n[s].replace("/>", ">")
		}
		if (l["!protect"]) n[n.length] = Vv(l["!protect"]);
		if (l["!autofilter"] != null) n[n.length] = Zv(l["!autofilter"], l, t, e);
		if (l["!merges"] != null && l["!merges"].length > 0) n[n.length] = Mv(l["!merges"]);
		var h = -1,
		d, v = -1;
		if (l["!links"].length > 0) {
			n[n.length] = "<hyperlinks>";
			l["!links"].forEach(function(e) {
				if (!e[1].Target) return;
				d = {
					ref: e[0]
				};
				if (e[1].Target.charAt(0) != "#") {
					v = vi(a, -1, lt(e[1].Target).replace(/#.*$/, ""), ci.HLINK);
					d["r:id"] = "rId" + v
				}
				if ((h = e[1].Target.indexOf("#")) > -1) d.location = lt(e[1].Target.slice(h + 1));
				if (e[1].Tooltip) d.tooltip = lt(e[1].Tooltip);
				d.display = e[1].display;
				n[n.length] = It("hyperlink", null, d)
			});
			n[n.length] = "</hyperlinks>"
		}
		delete l["!links"];
		if (l["!margins"] != null) n[n.length] = Gv(l["!margins"]);
		if (!r || r.ignoreEC || r.ignoreEC == void 0) n[n.length] = Rt("ignoredErrors", It("ignoredError", null, {
			numberStoredAsText: 1,
			sqref: o
		}));
		if (u.length > 0) {
			v = vi(a, -1, "../drawings/drawing" + (e + 1) + ".xml", ci.DRAW);
			n[n.length] = It("drawing", null, {
				"r:id": "rId" + v
			});
			l["!drawing"] = u
		}
		if (l["!comments"].length > 0) {
			v = vi(a, -1, "../drawings/vmlDrawing" + (e + 1) + ".vml", ci.VML);
			n[n.length] = It("legacyDrawing", null, {
				"r:id": "rId" + v
			});
			l["!legacy"] = v
		}
		if (n.length > 1) {
			n[n.length] = "</worksheet>";
			n[1] = n[1].replace("/>", ">")
		}
		return n.join("")
	}
	function np(e, r) {
		var t = {};
		var a = e.l + r;
		t.r = e._R(4);
		e.l += 4;
		var n = e._R(2);
		e.l += 1;
		var i = e._R(1);
		e.l = a;
		if (i & 7) t.level = i & 7;
		if (i & 16) t.hidden = true;
		if (i & 32) t.hpt = n / 20;
		return t
	}
	function ip(e, r, t) {
		var a = Ea(17 + 8 * 16);
		var n = (t["!rows"] || [])[e] || {};
		a._W(4, e);
		a._W(4, 0);
		var i = 320;
		if (n.hpx) i = fc(n.hpx) * 20;
		else if (n.hpt) i = n.hpt * 20;
		a._W(2, i);
		a._W(1, 0);
		var s = 0;
		if (n.level) s |= n.level;
		if (n.hidden) s |= 16;
		if (n.hpx || n.hpt) s |= 32;
		a._W(1, s);
		a._W(1, 0);
		var f = 0,
		l = a.l;
		a.l += 4;
		var o = {
			r: e,
			c: 0
		};
		var c = t["!data"] != null;
		for (var u = 0; u < 16; ++u) {
			if (r.s.c > u + 1 << 10 || r.e.c < u << 10) continue;
			var h = -1,
			d = -1;
			for (var v = u << 10; v < u + 1 << 10; ++v) {
				o.c = v;
				var p = c ? (t["!data"][o.r] || [])[o.c] : t[za(o)];
				if (p) {
					if (h < 0) h = v;
					d = v
				}
			}
			if (h < 0) continue; ++f;
			a._W(4, h);
			a._W(4, d)
		}
		var m = a.l;
		a.l = l;
		a._W(4, f);
		a.l = m;
		return a.length > a.l ? a.slice(0, a.l) : a
	}
	function sp(e, r, t, a) {
		var n = ip(a, t, r);
		if (n.length > 17 || (r["!rows"] || [])[a]) xa(e, 0, n)
	}
	var fp = Sn;
	var lp = xn;
	function op() {}
	function cp(e, r) {
		var t = {};
		var a = e[e.l]; ++e.l;
		t.above = !(a & 64);
		t.left = !(a & 128);
		e.l += 18;
		t.name = vn(e, r - 19);
		return t
	}
	function up(e, r, t) {
		if (t == null) t = Ea(84 + 4 * e.length);
		var a = 192;
		if (r) {
			if (r.above) a &= ~64;
			if (r.left) a &= ~128
		}
		t._W(1, a);
		for (var n = 1; n < 3; ++n) t._W(1, 0);
		On({
			auto: 1
		},
		t);
		t._W(-4, -1);
		t._W(-4, -1);
		pn(e, t);
		return t.slice(0, t.l)
	}
	function hp(e) {
		var r = cn(e);
		return [r]
	}
	function dp(e, r, t) {
		if (t == null) t = Ea(8);
		return un(r, t)
	}
	function vp(e) {
		var r = hn(e);
		return [r]
	}
	function pp(e, r, t) {
		if (t == null) t = Ea(4);
		return dn(r, t)
	}
	function mp(e) {
		var r = cn(e);
		var t = e._R(1);
		return [r, t, "b"]
	}
	function bp(e, r, t) {
		if (t == null) t = Ea(9);
		un(r, t);
		t._W(1, e.v ? 1 : 0);
		return t
	}
	function gp(e) {
		var r = hn(e);
		var t = e._R(1);
		return [r, t, "b"]
	}
	function wp(e, r, t) {
		if (t == null) t = Ea(5);
		dn(r, t);
		t._W(1, e.v ? 1 : 0);
		return t
	}
	function kp(e) {
		var r = cn(e);
		var t = e._R(1);
		return [r, t, "e"]
	}
	function Tp(e, r, t) {
		if (t == null) t = Ea(9);
		un(r, t);
		t._W(1, e.v);
		return t
	}
	function yp(e) {
		var r = hn(e);
		var t = e._R(1);
		return [r, t, "e"]
	}
	function Ep(e, r, t) {
		if (t == null) t = Ea(8);
		dn(r, t);
		t._W(1, e.v);
		t._W(2, 0);
		t._W(1, 0);
		return t
	}
	function _p(e) {
		var r = cn(e);
		var t = e._R(4);
		return [r, t, "s"]
	}
	function Sp(e, r, t) {
		if (t == null) t = Ea(12);
		un(r, t);
		t._W(4, r.v);
		return t
	}
	function xp(e) {
		var r = hn(e);
		var t = e._R(4);
		return [r, t, "s"]
	}
	function Ap(e, r, t) {
		if (t == null) t = Ea(8);
		dn(r, t);
		t._W(4, r.v);
		return t
	}
	function Cp(e) {
		var r = cn(e);
		var t = An(e);
		return [r, t, "n"]
	}
	function Rp(e, r, t) {
		if (t == null) t = Ea(16);
		un(r, t);
		Cn(e.v, t);
		return t
	}
	function Op(e) {
		var r = hn(e);
		var t = An(e);
		return [r, t, "n"]
	}
	function Ip(e, r, t) {
		if (t == null) t = Ea(12);
		dn(r, t);
		Cn(e.v, t);
		return t
	}
	function Np(e) {
		var r = cn(e);
		var t = Tn(e);
		return [r, t, "n"]
	}
	function Fp(e, r, t) {
		if (t == null) t = Ea(12);
		un(r, t);
		yn(e.v, t);
		return t
	}
	function Dp(e) {
		var r = hn(e);
		var t = Tn(e);
		return [r, t, "n"]
	}
	function Pp(e, r, t) {
		if (t == null) t = Ea(8);
		dn(r, t);
		yn(e.v, t);
		return t
	}
	function Lp(e) {
		var r = cn(e);
		var t = sn(e);
		return [r, t, "is"]
	}
	function Mp(e) {
		var r = cn(e);
		var t = rn(e);
		return [r, t, "str"]
	}
	function Up(e, r, t) {
		var a = e.v == null ? "": String(e.v);
		if (t == null) t = Ea(12 + 4 * e.v.length);
		un(r, t);
		tn(a, t);
		return t.length > t.l ? t.slice(0, t.l) : t
	}
	function Bp(e) {
		var r = hn(e);
		var t = rn(e);
		return [r, t, "str"]
	}
	function Wp(e, r, t) {
		var a = e.v == null ? "": String(e.v);
		if (t == null) t = Ea(8 + 4 * a.length);
		dn(r, t);
		tn(a, t);
		return t.length > t.l ? t.slice(0, t.l) : t
	}
	function zp(e, r, t) {
		var a = e.l + r;
		var n = cn(e);
		n.r = t["!row"];
		var i = e._R(1);
		var s = [n, i, "b"];
		if (t.cellFormula) {
			e.l += 2;
			var f = Kd(e, a - e.l, t);
			s[3] = Md(f, null, n, t.supbooks, t)
		} else e.l = a;
		return s
	}
	function Hp(e, r, t) {
		var a = e.l + r;
		var n = cn(e);
		n.r = t["!row"];
		var i = e._R(1);
		var s = [n, i, "e"];
		if (t.cellFormula) {
			e.l += 2;
			var f = Kd(e, a - e.l, t);
			s[3] = Md(f, null, n, t.supbooks, t)
		} else e.l = a;
		return s
	}
	function Vp(e, r, t) {
		var a = e.l + r;
		var n = cn(e);
		n.r = t["!row"];
		var i = An(e);
		var s = [n, i, "n"];
		if (t.cellFormula) {
			e.l += 2;
			var f = Kd(e, a - e.l, t);
			s[3] = Md(f, null, n, t.supbooks, t)
		} else e.l = a;
		return s
	}
	function $p(e, r, t) {
		var a = e.l + r;
		var n = cn(e);
		n.r = t["!row"];
		var i = rn(e);
		var s = [n, i, "str"];
		if (t.cellFormula) {
			e.l += 2;
			var f = Kd(e, a - e.l, t);
			s[3] = Md(f, null, n, t.supbooks, t)
		} else e.l = a;
		return s
	}
	var Xp = Sn;
	var Gp = xn;
	function jp(e, r) {
		if (r == null) r = Ea(4);
		r._W(4, e);
		return r
	}
	function Kp(e, r) {
		var t = e.l + r;
		var a = Sn(e, 16);
		var n = mn(e);
		var i = rn(e);
		var s = rn(e);
		var f = rn(e);
		e.l = t;
		var l = {
			rfx: a,
			relId: n,
			loc: i,
			display: f
		};
		if (s) l.Tooltip = s;
		return l
	}
	function Yp(e, r) {
		var t = Ea(50 + 4 * (e[1].Target.length + (e[1].Tooltip || "").length));
		xn({
			s: Wa(e[0]),
			e: Wa(e[0])
		},
		t);
		kn("rId" + r, t);
		var a = e[1].Target.indexOf("#");
		var n = a == -1 ? "": e[1].Target.slice(a + 1);
		tn(n || "", t);
		tn(e[1].Tooltip || "", t);
		tn("", t);
		return t.slice(0, t.l)
	}
	function Zp() {}
	function Jp(e, r, t) {
		var a = e.l + r;
		var n = En(e, 16);
		var i = e._R(1);
		var s = [n];
		s[2] = i;
		if (t.cellFormula) {
			var f = jd(e, a - e.l, t);
			s[1] = f
		} else e.l = a;
		return s
	}
	function qp(e, r, t) {
		var a = e.l + r;
		var n = Sn(e, 16);
		var i = [n];
		if (t.cellFormula) {
			var s = Zd(e, a - e.l, t);
			i[1] = s;
			e.l = a
		} else e.l = a;
		return i
	}
	function Qp(e, r, t) {
		if (t == null) t = Ea(18);
		var a = kv(e, r);
		t._W(-4, e);
		t._W(-4, e);
		t._W(4, (a.width || 10) * 256);
		t._W(4, 0);
		var n = 0;
		if (r.hidden) n |= 1;
		if (typeof a.width == "number") n |= 2;
		if (r.level) n |= r.level << 8;
		t._W(2, n);
		return t
	}
	var em = ["left", "right", "top", "bottom", "header", "footer"];
	function rm(e) {
		var r = {};
		em.forEach(function(t) {
			r[t] = An(e, 8)
		});
		return r
	}
	function tm(e, r) {
		if (r == null) r = Ea(6 * 8);
		Tv(e);
		em.forEach(function(t) {
			Cn(e[t], r)
		});
		return r
	}
	function am(e) {
		var r = e._R(2);
		e.l += 28;
		return {
			RTL: r & 32
		}
	}
	function nm(e, r, t) {
		if (t == null) t = Ea(30);
		var a = 924;
		if ((((r || {}).Views || [])[0] || {}).RTL) a |= 32;
		t._W(2, a);
		t._W(4, 0);
		t._W(4, 0);
		t._W(4, 0);
		t._W(1, 0);
		t._W(1, 0);
		t._W(2, 0);
		t._W(2, 100);
		t._W(2, 0);
		t._W(2, 0);
		t._W(2, 0);
		t._W(4, 0);
		return t
	}
	function im(e) {
		var r = Ea(24);
		r._W(4, 4);
		r._W(4, 1);
		xn(e, r);
		return r
	}
	function sm(e, r) {
		if (r == null) r = Ea(16 * 4 + 2);
		r._W(2, e.password ? Fo(e.password) : 0);
		r._W(4, 1); [["objects", false], ["scenarios", false], ["formatCells", true], ["formatColumns", true], ["formatRows", true], ["insertColumns", true], ["insertRows", true], ["insertHyperlinks", true], ["deleteColumns", true], ["deleteRows", true], ["selectLockedCells", false], ["sort", true], ["autoFilter", true], ["pivotTables", true], ["selectUnlockedCells", false]].forEach(function(t) {
			if (t[1]) r._W(4, e[t[0]] != null && !e[t[0]] ? 1 : 0);
			else r._W(4, e[t[0]] != null && e[t[0]] ? 0 : 1)
		});
		return r
	}
	function fm() {}
	function lm() {}
	function om(e, r, t, a, n, i, s) {
		if (!e) return e;
		var f = r || {};
		if (!a) a = {
			"!id": {}
		};
		if (g != null && f.dense == null) f.dense = g;
		var l = {};
		if (f.dense) l["!data"] = [];
		var o;
		var c = {
			s: {
				r: 2e6,
				c: 2e6
			},
			e: {
				r: 0,
				c: 0
			}
		};
		var u = [];
		var h = false,
		d = false;
		var v, p, m, b, w, k, T, y, E;
		var _ = [];
		f.biff = 12;
		f["!row"] = 0;
		var S = 0,
		x = false;
		var A = [];
		var C = {};
		var R = f.supbooks || n.supbooks || [[]];
		R.sharedf = C;
		R.arrayf = A;
		R.SheetNames = n.SheetNames || n.Sheets.map(function(e) {
			return e.name
		});
		if (!f.supbooks) {
			f.supbooks = R;
			if (n.Names) for (var O = 0; O < n.Names.length; ++O) R[0][O + 1] = n.Names[O]
		}
		var I = [],
		N = [];
		var F = false;
		Qb[16] = {
			n: "BrtShortReal",
			f: Op
		};
		var D, P;
		var L = 1462 * +!!((n || {}).WBProps || {}).date1904;
		_a(e,
		function U(e, r, g) {
			if (d) return;
			switch (g) {
			case 148:
				o = e;
				break;
			case 0:
				v = e;
				if (f.sheetRows && f.sheetRows <= v.r) d = true;
				y = Na(b = v.r);
				f["!row"] = v.r;
				if (e.hidden || e.hpt || e.level != null) {
					if (e.hpt) e.hpx = lc(e.hpt);
					N[e.r] = e
				}
				break;
			case 2:
				;
			case 3:
				;
			case 4:
				;
			case 5:
				;
			case 6:
				;
			case 7:
				;
			case 8:
				;
			case 9:
				;
			case 10:
				;
			case 11:
				;
			case 13:
				;
			case 14:
				;
			case 15:
				;
			case 16:
				;
			case 17:
				;
			case 18:
				;
			case 62:
				p = {
					t: e[2]
				};
				switch (e[2]) {
				case "n":
					p.v = e[1];
					break;
				case "s":
					T = mv[e[1]];
					p.v = T.t;
					p.r = T.r;
					break;
				case "b":
					p.v = e[1] ? true: false;
					break;
				case "e":
					p.v = e[1];
					if (f.cellText !== false) p.w = ti[p.v];
					break;
				case "str":
					p.t = "s";
					p.v = e[1];
					break;
				case "is":
					p.t = "s";
					p.v = e[1].t;
					break;
				}
				if (m = s.CellXf[e[0].iStyleRef]) Ev(p, m.numFmtId, null, f, i, s, L > 0);
				w = e[0].c == -1 ? w + 1 : e[0].c;
				if (f.dense) {
					if (!l["!data"][b]) l["!data"][b] = [];
					l["!data"][b][w] = p
				} else l[La(w) + y] = p;
				if (f.cellFormula) {
					x = false;
					for (S = 0; S < A.length; ++S) {
						var O = A[S];
						if (v.r >= O[0].s.r && v.r <= O[0].e.r) if (w >= O[0].s.c && w <= O[0].e.c) {
							p.F = Va(O[0]);
							x = true
						}
					}
					if (!x && e.length > 3) p.f = e[3]
				}
				if (c.s.r > v.r) c.s.r = v.r;
				if (c.s.c > w) c.s.c = w;
				if (c.e.r < v.r) c.e.r = v.r;
				if (c.e.c < w) c.e.c = w;
				if (f.cellDates && m && p.t == "n" && Le(q[m.numFmtId])) {
					var M = ae(p.v + L);
					if (M) {
						p.t = "d";
						p.v = new Date(Date.UTC(M.y, M.m - 1, M.d, M.H, M.M, M.S, M.u))
					}
				}
				if (D) {
					if (D.type == "XLDAPR") p.D = true;
					D = void 0
				}
				if (P) P = void 0;
				break;
			case 1:
				;
			case 12:
				if (!f.sheetStubs || h) break;
				p = {
					t: "z",
					v: void 0
				};
				w = e[0].c == -1 ? w + 1 : e[0].c;
				if (f.dense) {
					if (!l["!data"][b]) l["!data"][b] = [];
					l["!data"][b][w] = p
				} else l[La(w) + y] = p;
				if (c.s.r > v.r) c.s.r = v.r;
				if (c.s.c > w) c.s.c = w;
				if (c.e.r < v.r) c.e.r = v.r;
				if (c.e.c < w) c.e.c = w;
				if (D) {
					if (D.type == "XLDAPR") p.D = true;
					D = void 0
				}
				if (P) P = void 0;
				break;
			case 176:
				_.push(e);
				break;
			case 49:
				{
					D = ((f.xlmeta || {}).Cell || [])[e - 1]
				}
				break;
			case 494:
				var U = a["!id"][e.relId];
				if (U) {
					e.Target = U.Target;
					if (e.loc) e.Target += "#" + e.loc;
					e.Rel = U
				} else if (e.relId == "") {
					e.Target = "#" + e.loc
				}
				for (b = e.rfx.s.r; b <= e.rfx.e.r; ++b) for (w = e.rfx.s.c; w <= e.rfx.e.c; ++w) {
					if (f.dense) {
						if (!l["!data"][b]) l["!data"][b] = [];
						if (!l["!data"][b][w]) l["!data"][b][w] = {
							t: "z",
							v: undefined
						};
						l["!data"][b][w].l = e
					} else {
						k = La(w) + Na(b);
						if (!l[k]) l[k] = {
							t: "z",
							v: undefined
						};
						l[k].l = e
					}
				}
				break;
			case 426:
				if (!f.cellFormula) break;
				A.push(e);
				E = f.dense ? l["!data"][b][w] : l[La(w) + y];
				E.f = Md(e[1], c, {
					r: v.r,
					c: w
				},
				R, f);
				E.F = Va(e[0]);
				break;
			case 427:
				if (!f.cellFormula) break;
				C[za(e[0].s)] = e[1];
				E = f.dense ? l["!data"][b][w] : l[La(w) + y];
				E.f = Md(e[1], c, {
					r: v.r,
					c: w
				},
				R, f);
				break;
			case 60:
				if (!f.cellStyles) break;
				while (e.e >= e.s) {
					I[e.e--] = {
						width: e.w / 256,
						hidden: !!(e.flags & 1),
						level: e.level
					};
					if (!F) {
						F = true;
						ac(e.w / 256)
					}
					nc(I[e.e + 1])
				}
				break;
			case 551:
				if (e) l["!legrel"] = e;
				break;
			case 161:
				l["!autofilter"] = {
					ref: Va(e)
				};
				break;
			case 476:
				l["!margins"] = e;
				break;
			case 147:
				if (!n.Sheets[t]) n.Sheets[t] = {};
				if (e.name) n.Sheets[t].CodeName = e.name;
				if (e.above || e.left) l["!outline"] = {
					above: e.above,
					left: e.left
				};
				break;
			case 137:
				if (!n.Views) n.Views = [{}];
				if (!n.Views[0]) n.Views[0] = {};
				if (e.RTL) n.Views[0].RTL = true;
				break;
			case 485:
				break;
			case 64:
				;
			case 1053:
				break;
			case 151:
				break;
			case 152:
				;
			case 175:
				;
			case 644:
				;
			case 625:
				;
			case 562:
				;
			case 396:
				;
			case 1112:
				;
			case 1146:
				;
			case 471:
				;
			case 1050:
				;
			case 649:
				;
			case 1105:
				;
			case 589:
				;
			case 607:
				;
			case 564:
				;
			case 1055:
				;
			case 168:
				;
			case 174:
				;
			case 1180:
				;
			case 499:
				;
			case 507:
				;
			case 550:
				;
			case 171:
				;
			case 167:
				;
			case 1177:
				;
			case 169:
				;
			case 1181:
				;
			case 552:
				;
			case 661:
				;
			case 639:
				;
			case 478:
				;
			case 537:
				;
			case 477:
				;
			case 536:
				;
			case 1103:
				;
			case 680:
				;
			case 1104:
				;
			case 1024:
				;
			case 663:
				;
			case 535:
				;
			case 678:
				;
			case 504:
				;
			case 1043:
				;
			case 428:
				;
			case 170:
				;
			case 3072:
				;
			case 50:
				;
			case 2070:
				;
			case 1045:
				break;
			case 35:
				h = true;
				break;
			case 36:
				h = false;
				break;
			case 37:
				u.push(g);
				h = true;
				break;
			case 38:
				u.pop();
				h = false;
				break;
			default:
				if (r.T) {} else if (!h || f.WTF) throw new Error("Unexpected record 0x" + g.toString(16));
			}
		},
		f);
		delete f.supbooks;
		delete f["!row"];
		if (!l["!ref"] && (c.s.r < 2e6 || o && (o.e.r > 0 || o.e.c > 0 || o.s.r > 0 || o.s.c > 0))) l["!ref"] = Va(o || c);
		if (f.sheetRows && l["!ref"]) {
			var M = Ga(l["!ref"]);
			if (f.sheetRows <= +M.e.r) {
				M.e.r = f.sheetRows - 1;
				if (M.e.r > c.e.r) M.e.r = c.e.r;
				if (M.e.r < M.s.r) M.s.r = M.e.r;
				if (M.e.c > c.e.c) M.e.c = c.e.c;
				if (M.e.c < M.s.c) M.s.c = M.e.c;
				l["!fullref"] = l["!ref"];
				l["!ref"] = Va(M)
			}
		}
		if (_.length > 0) l["!merges"] = _;
		if (I.length > 0) l["!cols"] = I;
		if (N.length > 0) l["!rows"] = N;
		if (a["!id"][l["!legrel"]]) l["!legdrawel"] = a["!id"][l["!legrel"]];
		return l
	}
	function cm(e, r, t, a, n, i, s, f) {
		var l = {
			r: t,
			c: a
		};
		if (r.c) i["!comments"].push([za(l), r.c]);
		if (r.v === undefined) return false;
		var o = "";
		switch (r.t) {
		case "b":
			o = r.v ? "1": "0";
			break;
		case "d":
			r = Tr(r);
			r.z = r.z || q[14];
			r.v = dr(wr(r.v, f), f);
			r.t = "n";
			break;
		case "n":
			;
		case "e":
			o = "" + r.v;
			break;
		default:
			o = r.v;
			break;
		}
		l.s = yv(n.cellXfs, r, n);
		if (r.l) i["!links"].push([za(l), r.l]);
		switch (r.t) {
		case "s":
			;
		case "str":
			if (n.bookSST) {
				o = wv(n.Strings, r.v == null ? "": String(r.v), n.revStrings);
				l.t = "s";
				l.v = o;
				if (s) xa(e, 18, Ap(r, l));
				else xa(e, 7, Sp(r, l))
			} else {
				l.t = "str";
				if (s) xa(e, 17, Wp(r, l));
				else xa(e, 6, Up(r, l))
			}
			return true;
		case "n":
			if (r.v == (r.v | 0) && r.v > -1e3 && r.v < 1e3) {
				if (s) xa(e, 13, Pp(r, l));
				else xa(e, 2, Fp(r, l))
			} else if (isNaN(r.v)) {
				if (s) xa(e, 14, Ep({
					t: "e",
					v: 36
				},
				l));
				else xa(e, 3, Tp({
					t: "e",
					v: 36
				},
				l))
			} else if (!isFinite(r.v)) {
				if (s) xa(e, 14, Ep({
					t: "e",
					v: 7
				},
				l));
				else xa(e, 3, Tp({
					t: "e",
					v: 7
				},
				l))
			} else {
				if (s) xa(e, 16, Ip(r, l));
				else xa(e, 5, Rp(r, l))
			}
			return true;
		case "b":
			l.t = "b";
			if (s) xa(e, 15, wp(r, l));
			else xa(e, 4, bp(r, l));
			return true;
		case "e":
			l.t = "e";
			if (s) xa(e, 14, Ep(r, l));
			else xa(e, 3, Tp(r, l));
			return true;
		}
		if (s) xa(e, 12, pp(r, l));
		else xa(e, 1, dp(r, l));
		return true
	}
	function um(e, r, t, a, n) {
		var i = Ga(r["!ref"] || "A1"),
		s,
		f = "",
		l = [];
		var o = (((n || {}).Workbook || {}).WBProps || {}).date1904;
		xa(e, 145);
		var c = r["!data"] != null;
		var u = i.e.r;
		if (r["!rows"]) u = Math.max(i.e.r, r["!rows"].length - 1);
		for (var h = i.s.r; h <= u; ++h) {
			f = Na(h);
			sp(e, r, i, h);
			var d = false;
			if (h <= i.e.r) for (var v = i.s.c; v <= i.e.c; ++v) {
				if (h === i.s.r) l[v] = La(v);
				s = l[v] + f;
				var p = c ? (r["!data"][h] || [])[v] : r[s];
				if (!p) {
					d = false;
					continue
				}
				d = cm(e, p, h, v, a, r, d, o)
			}
		}
		xa(e, 146)
	}
	function hm(e, r) {
		if (!r || !r["!merges"]) return;
		xa(e, 177, jp(r["!merges"].length));
		r["!merges"].forEach(function(r) {
			xa(e, 176, Gp(r))
		});
		xa(e, 178)
	}
	function dm(e, r) {
		if (!r || !r["!cols"]) return;
		xa(e, 390);
		r["!cols"].forEach(function(r, t) {
			if (r) xa(e, 60, Qp(t, r))
		});
		xa(e, 391)
	}
	function vm(e, r) {
		if (!r || !r["!ref"]) return;
		xa(e, 648);
		xa(e, 649, im(Ga(r["!ref"])));
		xa(e, 650)
	}
	function pm(e, r, t) {
		r["!links"].forEach(function(r) {
			if (!r[1].Target) return;
			var a = vi(t, -1, r[1].Target.replace(/#.*$/, ""), ci.HLINK);
			xa(e, 494, Yp(r, a))
		});
		delete r["!links"]
	}
	function mm(e, r, t, a) {
		if (r["!comments"].length > 0) {
			var n = vi(a, -1, "../drawings/vmlDrawing" + (t + 1) + ".vml", ci.VML);
			xa(e, 551, kn("rId" + n));
			r["!legacy"] = n
		}
	}
	function bm(e, r, t, a) {
		if (!r["!autofilter"]) return;
		var n = r["!autofilter"];
		var i = typeof n.ref === "string" ? n.ref: Va(n.ref);
		if (!t.Workbook) t.Workbook = {
			Sheets: []
		};
		if (!t.Workbook.Names) t.Workbook.Names = [];
		var s = t.Workbook.Names;
		var f = Ha(i);
		if (f.s.r == f.e.r) {
			f.e.r = Ha(r["!ref"]).e.r;
			i = Va(f)
		}
		for (var l = 0; l < s.length; ++l) {
			var o = s[l];
			if (o.Name != "_xlnm._FilterDatabase") continue;
			if (o.Sheet != a) continue;
			o.Ref = Xa(t.SheetNames[a]) + "!" + $a(i);
			break
		}
		if (l == s.length) s.push({
			Name: "_xlnm._FilterDatabase",
			Sheet: a,
			Ref: Xa(t.SheetNames[a]) + "!" + $a(i)
		});
		xa(e, 161, xn(Ga(i)));
		xa(e, 162)
	}
	function gm(e, r, t) {
		xa(e, 133); {
			xa(e, 137, nm(r, t));
			xa(e, 138)
		}
		xa(e, 134)
	}
	function wm() {}
	function km(e, r) {
		if (!r["!protect"]) return;
		xa(e, 535, sm(r["!protect"]))
	}
	function Tm(e, r, t, a) {
		var n = Sa();
		var i = t.SheetNames[e],
		s = t.Sheets[i] || {};
		var f = i;
		try {
			if (t && t.Workbook) f = t.Workbook.Sheets[e].CodeName || f
		} catch(l) {}
		var o = Ga(s["!ref"] || "A1");
		if (o.e.c > 16383 || o.e.r > 1048575) {
			if (r.WTF) throw new Error("Range " + (s["!ref"] || "A1") + " exceeds format limit A1:XFD1048576");
			o.e.c = Math.min(o.e.c, 16383);
			o.e.r = Math.min(o.e.c, 1048575)
		}
		s["!links"] = [];
		s["!comments"] = [];
		xa(n, 129);
		if (t.vbaraw || s["!outline"]) xa(n, 147, up(f, s["!outline"]));
		xa(n, 148, lp(o));
		gm(n, s, t.Workbook);
		wm(n, s);
		dm(n, s, e, r, t);
		um(n, s, e, r, t);
		km(n, s);
		bm(n, s, t, e);
		hm(n, s);
		pm(n, s, a);
		if (s["!margins"]) xa(n, 476, tm(s["!margins"]));
		if (!r || r.ignoreEC || r.ignoreEC == void 0) vm(n, s);
		mm(n, s, e, a);
		xa(n, 130);
		return n.end()
	}
	function ym(e) {
		var r = [];
		var t = e.match(/^<c:numCache>/);
		var a; (e.match(/<c:pt idx="(\d*)">(.*?)<\/c:pt>/gm) || []).forEach(function(e) {
			var a = e.match(/<c:pt idx="(\d*?)"><c:v>(.*)<\/c:v><\/c:pt>/);
			if (!a) return;
			r[ + a[1]] = t ? +a[2] : a[2]
		});
		var n = it((e.match(/<c:formatCode>([\s\S]*?)<\/c:formatCode>/) || ["", "General"])[1]); (e.match(/<c:f>(.*?)<\/c:f>/gm) || []).forEach(function(e) {
			a = e.replace(/<.*?>/g, "")
		});
		return [r, n, a]
	}
	function Em(e, r, t, a, n, i) {
		var s = i || {
			"!type": "chart"
		};
		if (!e) return i;
		var f = 0,
		l = 0,
		o = "A";
		var c = {
			s: {
				r: 2e6,
				c: 2e6
			},
			e: {
				r: 0,
				c: 0
			}
		}; (e.match(/<c:numCache>[\s\S]*?<\/c:numCache>/gm) || []).forEach(function(e) {
			var r = ym(e);
			c.s.r = c.s.c = 0;
			c.e.c = f;
			o = La(f);
			r[0].forEach(function(e, t) {
				if (s["!data"]) {
					if (!s["!data"][t]) s["!data"][t] = [];
					s["!data"][t][f] = {
						t: "n",
						v: e,
						z: r[1]
					}
				} else s[o + Na(t)] = {
					t: "n",
					v: e,
					z: r[1]
				};
				l = t
			});
			if (c.e.r < l) c.e.r = l; ++f
		});
		if (f > 0) s["!ref"] = Va(c);
		return s
	}
	function _m(e, r, t, a, n) {
		if (!e) return e;
		if (!a) a = {
			"!id": {}
		};
		var i = {
			"!type": "chart",
			"!drawel": null,
			"!rel": ""
		};
		var s;
		var f = e.match(Fv);
		if (f) Uv(f[0], i, n, t);
		if (s = e.match(/drawing r:id="(.*?)"/)) i["!rel"] = s[1];
		if (a["!id"][i["!rel"]]) i["!drawel"] = a["!id"][i["!rel"]];
		return i
	}
	function Sm(e, r) {
		e.l += 10;
		var t = rn(e, r - 10);
		return {
			name: t
		}
	}
	function xm(e, r, t, a, n) {
		if (!e) return e;
		if (!a) a = {
			"!id": {}
		};
		var i = {
			"!type": "chart",
			"!drawel": null,
			"!rel": ""
		};
		var s = [];
		var f = false;
		_a(e,
		function l(e, a, o) {
			switch (o) {
			case 550:
				i["!rel"] = e;
				break;
			case 651:
				if (!n.Sheets[t]) n.Sheets[t] = {};
				if (e.name) n.Sheets[t].CodeName = e.name;
				break;
			case 562:
				;
			case 652:
				;
			case 669:
				;
			case 679:
				;
			case 551:
				;
			case 552:
				;
			case 476:
				;
			case 3072:
				break;
			case 35:
				f = true;
				break;
			case 36:
				f = false;
				break;
			case 37:
				s.push(o);
				break;
			case 38:
				s.pop();
				break;
			default:
				if (a.T > 0) s.push(o);
				else if (a.T < 0) s.pop();
				else if (!f || r.WTF) throw new Error("Unexpected record 0x" + o.toString(16));
			}
		},
		r);
		if (a["!id"][i["!rel"]]) i["!drawel"] = a["!id"][i["!rel"]];
		return i
	}
	var Am = [["allowRefreshQuery", false, "bool"], ["autoCompressPictures", true, "bool"], ["backupFile", false, "bool"], ["checkCompatibility", false, "bool"], ["CodeName", ""], ["date1904", false, "bool"], ["defaultThemeVersion", 0, "int"], ["filterPrivacy", false, "bool"], ["hidePivotFieldList", false, "bool"], ["promptedSolutions", false, "bool"], ["publishItems", false, "bool"], ["refreshAllConnections", false, "bool"], ["saveExternalLinkValues", true, "bool"], ["showBorderUnselectedTables", true, "bool"], ["showInkAnnotation", true, "bool"], ["showObjects", "all"], ["showPivotChartFilter", false, "bool"], ["updateLinks", "userSet"]];
	var Cm = [["activeTab", 0, "int"], ["autoFilterDateGrouping", true, "bool"], ["firstSheet", 0, "int"], ["minimized", false, "bool"], ["showHorizontalScroll", true, "bool"], ["showSheetTabs", true, "bool"], ["showVerticalScroll", true, "bool"], ["tabRatio", 600, "int"], ["visibility", "visible"]];
	var Rm = [];
	var Om = [["calcCompleted", "true"], ["calcMode", "auto"], ["calcOnSave", "true"], ["concurrentCalc", "true"], ["fullCalcOnLoad", "false"], ["fullPrecision", "true"], ["iterate", "false"], ["iterateCount", "100"], ["iterateDelta", "0.001"], ["refMode", "A1"]];
	function Im(e, r) {
		for (var t = 0; t != e.length; ++t) {
			var a = e[t];
			for (var n = 0; n != r.length; ++n) {
				var i = r[n];
				if (a[i[0]] == null) a[i[0]] = i[1];
				else switch (i[2]) {
				case "bool":
					if (typeof a[i[0]] == "string") a[i[0]] = pt(a[i[0]]);
					break;
				case "int":
					if (typeof a[i[0]] == "string") a[i[0]] = parseInt(a[i[0]], 10);
					break;
				}
			}
		}
	}
	function Nm(e, r) {
		for (var t = 0; t != r.length; ++t) {
			var a = r[t];
			if (e[a[0]] == null) e[a[0]] = a[1];
			else switch (a[2]) {
			case "bool":
				if (typeof e[a[0]] == "string") e[a[0]] = pt(e[a[0]]);
				break;
			case "int":
				if (typeof e[a[0]] == "string") e[a[0]] = parseInt(e[a[0]], 10);
				break;
			}
		}
	}
	function Fm(e) {
		Nm(e.WBProps, Am);
		Nm(e.CalcPr, Om);
		Im(e.WBView, Cm);
		Im(e.Sheets, Rm);
		bv.date1904 = pt(e.WBProps.date1904)
	}
	function Dm(e) {
		if (!e.Workbook) return "false";
		if (!e.Workbook.WBProps) return "false";
		return pt(e.Workbook.WBProps.date1904) ? "true": "false"
	}
	var Pm = ":][*?/\\".split("");
	function Lm(e, r) {
		try {
			if (e == "") throw new Error("Sheet name cannot be blank");
			if (e.length > 31) throw new Error("Sheet name cannot exceed 31 chars");
			if (e.charCodeAt(0) == 39 || e.charCodeAt(e.length - 1) == 39) throw new Error("Sheet name cannot start or end with apostrophe (')");
			if (e.toLowerCase() == "history") throw new Error("Sheet name cannot be 'History'");
			Pm.forEach(function(r) {
				if (e.indexOf(r) == -1) return;
				throw new Error("Sheet name cannot contain : \\ / ? * [ ]")
			})
		} catch(t) {
			if (r) return false;
			throw t
		}
		return true
	}
	function Mm(e, r, t) {
		e.forEach(function(a, n) {
			Lm(a);
			for (var i = 0; i < n; ++i) if (a == e[i]) throw new Error("Duplicate Sheet Name: " + a);
			if (t) {
				var s = r && r[n] && r[n].CodeName || a;
				if (s.charCodeAt(0) == 95 && s.length > 22) throw new Error("Bad Code Name: Worksheet" + s)
			}
		})
	}
	function Um(e) {
		if (!e || !e.SheetNames || !e.Sheets) throw new Error("Invalid Workbook");
		if (!e.SheetNames.length) throw new Error("Workbook is empty");
		var r = e.Workbook && e.Workbook.Sheets || [];
		Mm(e.SheetNames, r, !!e.vbaraw);
		for (var t = 0; t < e.SheetNames.length; ++t) _v(e.Sheets[e.SheetNames[t]], e.SheetNames[t], t);
		e.SheetNames.forEach(function(r, t) {
			var a = e.Sheets[r];
			if (!a || !a["!autofilter"]) return;
			var n;
			if (!e.Workbook) e.Workbook = {};
			if (!e.Workbook.Names) e.Workbook.Names = [];
			e.Workbook.Names.forEach(function(e) {
				if (e.Name == "_xlnm._FilterDatabase" && e.Sheet == t) n = e
			});
			var i = Xa(r) + "!" + $a(a["!autofilter"].ref);
			if (n) n.Ref = i;
			else e.Workbook.Names.push({
				Name: "_xlnm._FilterDatabase",
				Sheet: t,
				Ref: i
			})
		})
	}
	var Bm = /<\w+:workbook/;
	function Wm(e, r) {
		if (!e) throw new Error("Could not find file");
		var t = {
			AppVersion: {},
			WBProps: {},
			WBView: [],
			Sheets: [],
			CalcPr: {},
			Names: [],
			xmlns: ""
		};
		var a = false,
		n = "xmlns";
		var i = {},
		s = 0;
		e.replace(qr,
		function f(l, o) {
			var c = rt(l);
			switch (tt(c[0])) {
			case "<?xml":
				break;
			case "<workbook":
				if (l.match(Bm)) n = "xmlns" + l.match(/<(\w+):/)[1];
				t.xmlns = c[n];
				break;
			case "</workbook>":
				break;
			case "<fileVersion":
				delete c[0];
				t.AppVersion = c;
				break;
			case "<fileVersion/>":
				;
			case "</fileVersion>":
				break;
			case "<fileSharing":
				break;
			case "<fileSharing/>":
				break;
			case "<workbookPr":
				;
			case "<workbookPr/>":
				Am.forEach(function(e) {
					if (c[e[0]] == null) return;
					switch (e[2]) {
					case "bool":
						t.WBProps[e[0]] = pt(c[e[0]]);
						break;
					case "int":
						t.WBProps[e[0]] = parseInt(c[e[0]], 10);
						break;
					default:
						t.WBProps[e[0]] = c[e[0]];
					}
				});
				if (c.codeName) t.WBProps.CodeName = kt(c.codeName);
				break;
			case "</workbookPr>":
				break;
			case "<workbookProtection":
				break;
			case "<workbookProtection/>":
				break;
			case "<bookViews":
				;
			case "<bookViews>":
				;
			case "</bookViews>":
				break;
			case "<workbookView":
				;
			case "<workbookView/>":
				delete c[0];
				t.WBView.push(c);
				break;
			case "</workbookView>":
				break;
			case "<sheets":
				;
			case "<sheets>":
				;
			case "</sheets>":
				break;
			case "<sheet":
				switch (c.state) {
				case "hidden":
					c.Hidden = 1;
					break;
				case "veryHidden":
					c.Hidden = 2;
					break;
				default:
					c.Hidden = 0;
				}
				delete c.state;
				c.name = it(kt(c.name));
				delete c[0];
				t.Sheets.push(c);
				break;
			case "</sheet>":
				break;
			case "<functionGroups":
				;
			case "<functionGroups/>":
				break;
			case "<functionGroup":
				break;
			case "<externalReferences":
				;
			case "</externalReferences>":
				;
			case "<externalReferences>":
				break;
			case "<externalReference":
				break;
			case "<definedNames/>":
				break;
			case "<definedNames>":
				;
			case "<definedNames":
				a = true;
				break;
			case "</definedNames>":
				a = false;
				break;
			case "<definedName":
				{
					i = {};
					i.Name = kt(c.name);
					if (c.comment) i.Comment = c.comment;
					if (c.localSheetId) i.Sheet = +c.localSheetId;
					if (pt(c.hidden || "0")) i.Hidden = true;
					s = o + l.length
				}
				break;
			case "</definedName>":
				{
					i.Ref = it(kt(e.slice(s, o)));
					t.Names.push(i)
				}
				break;
			case "<definedName/>":
				break;
			case "<calcPr":
				delete c[0];
				t.CalcPr = c;
				break;
			case "<calcPr/>":
				delete c[0];
				t.CalcPr = c;
				break;
			case "</calcPr>":
				break;
			case "<oleSize":
				break;
			case "<customWorkbookViews>":
				;
			case "</customWorkbookViews>":
				;
			case "<customWorkbookViews":
				break;
			case "<customWorkbookView":
				;
			case "</customWorkbookView>":
				break;
			case "<pivotCaches>":
				;
			case "</pivotCaches>":
				;
			case "<pivotCaches":
				break;
			case "<pivotCache":
				break;
			case "<smartTagPr":
				;
			case "<smartTagPr/>":
				break;
			case "<smartTagTypes":
				;
			case "<smartTagTypes>":
				;
			case "</smartTagTypes>":
				break;
			case "<smartTagType":
				break;
			case "<webPublishing":
				;
			case "<webPublishing/>":
				break;
			case "<fileRecoveryPr":
				;
			case "<fileRecoveryPr/>":
				break;
			case "<webPublishObjects>":
				;
			case "<webPublishObjects":
				;
			case "</webPublishObjects>":
				break;
			case "<webPublishObject":
				break;
			case "<extLst":
				;
			case "<extLst>":
				;
			case "</extLst>":
				;
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
				;
			case "<AlternateContent>":
				a = true;
				break;
			case "</AlternateContent>":
				a = false;
				break;
			case "<revisionPtr":
				break;
			default:
				if (!a && r.WTF) throw new Error("unrecognized " + c[0] + " in workbook");
			}
			return l
		});
		if (Mt.indexOf(t.xmlns) === -1) throw new Error("Unknown Namespace: " + t.xmlns);
		Fm(t);
		return t
	}
	function zm(e) {
		var r = [Kr];
		r[r.length] = It("workbook", null, {
			xmlns: Mt[0],
			"xmlns:r": Lt.r
		});
		var t = e.Workbook && (e.Workbook.Names || []).length > 0;
		var a = {
			codeName: "ThisWorkbook"
		};
		if (e.Workbook && e.Workbook.WBProps) {
			Am.forEach(function(r) {
				if (e.Workbook.WBProps[r[0]] == null) return;
				if (e.Workbook.WBProps[r[0]] == r[1]) return;
				a[r[0]] = e.Workbook.WBProps[r[0]]
			});
			if (e.Workbook.WBProps.CodeName) {
				a.codeName = e.Workbook.WBProps.CodeName;
				delete a.CodeName
			}
		}
		r[r.length] = It("workbookPr", null, a);
		var n = e.Workbook && e.Workbook.Sheets || [];
		var i = 0;
		if (n && n[0] && !!n[0].Hidden) {
			r[r.length] = "<bookViews>";
			for (i = 0; i != e.SheetNames.length; ++i) {
				if (!n[i]) break;
				if (!n[i].Hidden) break
			}
			if (i == e.SheetNames.length) i = 0;
			r[r.length] = '<workbookView firstSheet="' + i + '" activeTab="' + i + '"/>';
			r[r.length] = "</bookViews>"
		}
		r[r.length] = "<sheets>";
		for (i = 0; i != e.SheetNames.length; ++i) {
			var s = {
				name: lt(e.SheetNames[i].slice(0, 31))
			};
			s.sheetId = "" + (i + 1);
			s["r:id"] = "rId" + (i + 1);
			if (n[i]) switch (n[i].Hidden) {
			case 1:
				s.state = "hidden";
				break;
			case 2:
				s.state = "veryHidden";
				break;
			}
			r[r.length] = It("sheet", null, s)
		}
		r[r.length] = "</sheets>";
		if (t) {
			r[r.length] = "<definedNames>";
			if (e.Workbook && e.Workbook.Names) e.Workbook.Names.forEach(function(e) {
				var t = {
					name: e.Name
				};
				if (e.Comment) t.comment = e.Comment;
				if (e.Sheet != null) t.localSheetId = "" + e.Sheet;
				if (e.Hidden) t.hidden = "1";
				if (!e.Ref) return;
				r[r.length] = It("definedName", lt(e.Ref), t)
			});
			r[r.length] = "</definedNames>"
		}
		if (r.length > 2) {
			r[r.length] = "</workbook>";
			r[1] = r[1].replace("/>", ">")
		}
		return r.join("")
	}
	function Hm(e, r) {
		var t = {};
		t.Hidden = e._R(4);
		t.iTabID = e._R(4);
		t.strRelID = wn(e, r - 8);
		t.name = rn(e);
		return t
	}
	function Vm(e, r) {
		if (!r) r = Ea(127);
		r._W(4, e.Hidden);
		r._W(4, e.iTabID);
		kn(e.strRelID, r);
		tn(e.name.slice(0, 31), r);
		return r.length > r.l ? r.slice(0, r.l) : r
	}
	function $m(e, r) {
		var t = {};
		var a = e._R(4);
		t.defaultThemeVersion = e._R(4);
		var n = r > 8 ? rn(e) : "";
		if (n.length > 0) t.CodeName = n;
		t.autoCompressPictures = !!(a & 65536);
		t.backupFile = !!(a & 64);
		t.checkCompatibility = !!(a & 4096);
		t.date1904 = !!(a & 1);
		t.filterPrivacy = !!(a & 8);
		t.hidePivotFieldList = !!(a & 1024);
		t.promptedSolutions = !!(a & 16);
		t.publishItems = !!(a & 2048);
		t.refreshAllConnections = !!(a & 262144);
		t.saveExternalLinkValues = !!(a & 128);
		t.showBorderUnselectedTables = !!(a & 4);
		t.showInkAnnotation = !!(a & 32);
		t.showObjects = ["all", "placeholders", "none"][a >> 13 & 3];
		t.showPivotChartFilter = !!(a & 32768);
		t.updateLinks = ["userSet", "never", "always"][a >> 8 & 3];
		return t
	}
	function Xm(e, r) {
		if (!r) r = Ea(72);
		var t = 0;
		if (e) {
			if (e.date1904) t |= 1;
			if (e.filterPrivacy) t |= 8
		}
		r._W(4, t);
		r._W(4, 0);
		pn(e && e.CodeName || "ThisWorkbook", r);
		return r.slice(0, r.l)
	}
	function Gm(e, r) {
		var t = {};
		e._R(4);
		t.ArchID = e._R(4);
		e.l += r - 8;
		return t
	}
	function jm(e, r, t) {
		var a = e.l + r;
		var n = e._R(4);
		e.l += 1;
		var i = e._R(4);
		var s = gn(e);
		var f = Yd(e, 0, t);
		var l = mn(e);
		if (n & 32) s = "_xlnm." + s;
		e.l = a;
		var o = {
			Name: s,
			Ptg: f,
			Flags: n
		};
		if (i < 268435455) o.Sheet = i;
		if (l) o.Comment = l;
		return o
	}
	function Km(e, r) {
		var t = Ea(9);
		var a = 0;
		var n = e.Name;
		if (ni.indexOf(n) > -1) {
			a |= 32;
			n = n.slice(6)
		}
		t._W(4, a);
		t._W(1, 0);
		t._W(4, e.Sheet == null ? 4294967295 : e.Sheet);
		var i = [t, tn(n), lv(e.Ref, r)];
		if (e.Comment) i.push(bn(e.Comment));
		else {
			var s = Ea(4);
			s._W(4, 4294967295);
			i.push(s)
		}
		return P(i)
	}
	function Ym(e, r) {
		var t = {
			AppVersion: {},
			WBProps: {},
			WBView: [],
			Sheets: [],
			CalcPr: {},
			xmlns: ""
		};
		var a = [];
		var n = false;
		if (!r) r = {};
		r.biff = 12;
		var i = [];
		var s = [[]];
		s.SheetNames = [];
		s.XTI = [];
		Qb[16] = {
			n: "BrtFRTArchID$",
			f: Gm
		};
		_a(e,
		function f(e, l, o) {
			switch (o) {
			case 156:
				s.SheetNames.push(e.name);
				t.Sheets.push(e);
				break;
			case 153:
				t.WBProps = e;
				break;
			case 39:
				if (e.Sheet != null) r.SID = e.Sheet;
				e.Ref = Md(e.Ptg, null, null, s, r);
				delete r.SID;
				delete e.Ptg;
				i.push(e);
				break;
			case 1036:
				break;
			case 357:
				;
			case 358:
				;
			case 355:
				;
			case 667:
				if (!s[0].length) s[0] = [o, e];
				else s.push([o, e]);
				s[s.length - 1].XTI = [];
				break;
			case 362:
				if (s.length === 0) {
					s[0] = [];
					s[0].XTI = []
				}
				s[s.length - 1].XTI = s[s.length - 1].XTI.concat(e);
				s.XTI = s.XTI.concat(e);
				break;
			case 361:
				break;
			case 2071:
				;
			case 158:
				;
			case 143:
				;
			case 664:
				;
			case 353:
				break;
			case 3072:
				;
			case 3073:
				;
			case 534:
				;
			case 677:
				;
			case 157:
				;
			case 610:
				;
			case 2050:
				;
			case 155:
				;
			case 548:
				;
			case 676:
				;
			case 128:
				;
			case 665:
				;
			case 2128:
				;
			case 2125:
				;
			case 549:
				;
			case 2053:
				;
			case 596:
				;
			case 2076:
				;
			case 2075:
				;
			case 2082:
				;
			case 397:
				;
			case 154:
				;
			case 1117:
				;
			case 553:
				;
			case 2091:
				break;
			case 35:
				a.push(o);
				n = true;
				break;
			case 36:
				a.pop();
				n = false;
				break;
			case 37:
				a.push(o);
				n = true;
				break;
			case 38:
				a.pop();
				n = false;
				break;
			case 16:
				break;
			default:
				if (l.T) {} else if (!n || r.WTF && a[a.length - 1] != 37 && a[a.length - 1] != 35) throw new Error("Unexpected record 0x" + o.toString(16));
			}
		},
		r);
		Fm(t);
		t.Names = i;
		t.supbooks = s;
		return t
	}
	function Zm(e, r) {
		xa(e, 143);
		for (var t = 0; t != r.SheetNames.length; ++t) {
			var a = r.Workbook && r.Workbook.Sheets && r.Workbook.Sheets[t] && r.Workbook.Sheets[t].Hidden || 0;
			var n = {
				Hidden: a,
				iTabID: t + 1,
				strRelID: "rId" + (t + 1),
				name: r.SheetNames[t]
			};
			xa(e, 156, Vm(n))
		}
		xa(e, 144)
	}
	function Jm(r, t) {
		if (!t) t = Ea(127);
		for (var a = 0; a != 4; ++a) t._W(4, 0);
		tn("SheetJS", t);
		tn(e.version, t);
		tn(e.version, t);
		tn("7262", t);
		return t.length > t.l ? t.slice(0, t.l) : t
	}
	function qm(e, r) {
		if (!r) r = Ea(29);
		r._W(-4, 0);
		r._W(-4, 460);
		r._W(4, 28800);
		r._W(4, 17600);
		r._W(4, 500);
		r._W(4, e);
		r._W(4, e);
		var t = 120;
		r._W(1, t);
		return r.length > r.l ? r.slice(0, r.l) : r
	}
	function Qm(e, r) {
		if (!r.Workbook || !r.Workbook.Sheets) return;
		var t = r.Workbook.Sheets;
		var a = 0,
		n = -1,
		i = -1;
		for (; a < t.length; ++a) {
			if (!t[a] || !t[a].Hidden && n == -1) n = a;
			else if (t[a].Hidden == 1 && i == -1) i = a
		}
		if (i > n) return;
		xa(e, 135);
		xa(e, 158, qm(n));
		xa(e, 136)
	}
	function eb(e, r) {
		if (!r.Workbook || !r.Workbook.Names) return;
		r.Workbook.Names.forEach(function(t) {
			try {
				if (t.Flags & 14) return;
				xa(e, 39, Km(t, r))
			} catch(a) {
				console.error("Could not serialize defined name " + JSON.stringify(t))
			}
		})
	}
	function rb(e) {
		var r = e.SheetNames.length;
		var t = Ea(12 * r + 28);
		t._W(4, r + 2);
		t._W(4, 0);
		t._W(4, -2);
		t._W(4, -2);
		t._W(4, 0);
		t._W(4, -1);
		t._W(4, -1);
		for (var a = 0; a < r; ++a) {
			t._W(4, 0);
			t._W(4, a);
			t._W(4, a)
		}
		return t
	}
	function tb(e, r) {
		xa(e, 353);
		xa(e, 357);
		xa(e, 362, rb(r, 0));
		xa(e, 354)
	}
	function ab(e, r) {
		var t = Sa();
		xa(t, 131);
		xa(t, 128, Jm());
		xa(t, 153, Xm(e.Workbook && e.Workbook.WBProps || null));
		Qm(t, e, r);
		Zm(t, e, r);
		tb(t, e);
		if ((e.Workbook || {}).Names) eb(t, e);
		xa(t, 132);
		return t.end()
	}
	function nb(e, r, t) {
		if (r.slice(-4) === ".bin") return Ym(e, t);
		return Wm(e, t)
	}
	function ib(e, r, t, a, n, i, s, f) {
		if (r.slice(-4) === ".bin") return om(e, a, t, n, i, s, f);
		return Lv(e, a, t, n, i, s, f)
	}
	function sb(e, r, t, a, n, i, s, f) {
		if (r.slice(-4) === ".bin") return xm(e, a, t, n, i, s, f);
		return _m(e, a, t, n, i, s, f)
	}
	function fb(e, r, t, a, n, i, s, f) {
		if (r.slice(-4) === ".bin") return Qu(e, a, t, n, i, s, f);
		return eh(e, a, t, n, i, s, f)
	}
	function lb(e, r, t, a, n, i, s, f) {
		if (r.slice(-4) === ".bin") return Ju(e, a, t, n, i, s, f);
		return qu(e, a, t, n, i, s, f)
	}
	function ob(e, r, t, a) {
		if (r.slice(-4) === ".bin") return Lc(e, t, a);
		return wc(e, t, a)
	}
	function cb(e, r, t) {
		if (r.slice(-4) === ".bin") return ho(e, t);
		return lo(e, t)
	}
	function ub(e, r, t) {
		if (r.slice(-4) === ".bin") return Xu(e, t);
		return Pu(e, t)
	}
	function hb(e, r, t) {
		if (r.slice(-4) === ".bin") return xu(e, r, t);
		return _u(e, r, t)
	}
	function db(e, r, t, a) {
		if (t.slice(-4) === ".bin") return Cu(e, r, t, a);
		return Au(e, r, t, a)
	}
	function vb(e, r, t) {
		if (r.slice(-4) === ".bin") return ku(e, r, t);
		return yu(e, r, t)
	}
	var pb = /([\w:]+)=((?:")([^"]*)(?:")|(?:')([^']*)(?:'))/g;
	var mb = /([\w:]+)=((?:")(?:[^"]*)(?:")|(?:')(?:[^']*)(?:'))/;
	function bb(e, r) {
		var t = e.split(/\s+/);
		var a = [];
		if (!r) a[0] = t[0];
		if (t.length === 1) return a;
		var n = e.match(pb),
		i,
		s,
		f,
		l;
		if (n) for (l = 0; l != n.length; ++l) {
			i = n[l].match(mb);
			if ((s = i[1].indexOf(":")) === -1) a[i[1]] = i[2].slice(1, i[2].length - 1);
			else {
				if (i[1].slice(0, 6) === "xmlns:") f = "xmlns" + i[1].slice(6);
				else f = i[1].slice(s + 1);
				a[f] = i[2].slice(1, i[2].length - 1)
			}
		}
		return a
	}
	function gb(e) {
		var r = e.split(/\s+/);
		var t = {};
		if (r.length === 1) return t;
		var a = e.match(pb),
		n,
		i,
		s,
		f;
		if (a) for (f = 0; f != a.length; ++f) {
			n = a[f].match(mb);
			if ((i = n[1].indexOf(":")) === -1) t[n[1]] = n[2].slice(1, n[2].length - 1);
			else {
				if (n[1].slice(0, 6) === "xmlns:") s = "xmlns" + n[1].slice(6);
				else s = n[1].slice(i + 1);
				t[s] = n[2].slice(1, n[2].length - 1)
			}
		}
		return t
	}
	var wb;
	function kb(e, r, t) {
		var a = wb[e] || it(e);
		if (a === "General") return oe(r);
		return ze(a, r, {
			date1904: !!t
		})
	}
	function Tb(e, r, t, a) {
		var n = a;
		switch ((t[0].match(/dt:dt="([\w.]+)"/) || ["", ""])[1]) {
		case "boolean":
			n = pt(a);
			break;
		case "i2":
			;
		case "int":
			n = parseInt(a, 10);
			break;
		case "r4":
			;
		case "float":
			n = parseFloat(a);
			break;
		case "date":
			;
		case "dateTime.tz":
			n = wr(a);
			break;
		case "i8":
			;
		case "string":
			;
		case "fixed":
			;
		case "uuid":
			;
		case "bin.base64":
			break;
		default:
			throw new Error("bad custprop:" + t[0]);
		}
		e[it(r)] = n
	}
	function yb(e, r, t, a) {
		if (e.t === "z") return;
		if (!t || t.cellText !== false) try {
			if (e.t === "e") {
				e.w = e.w || ti[e.v]
			} else if (r === "General") {
				if (e.t === "n") {
					if ((e.v | 0) === e.v) e.w = e.v.toString(10);
					else e.w = le(e.v)
				} else e.w = oe(e.v)
			} else e.w = kb(r || "General", e.v, a)
		} catch(n) {
			if (t.WTF) throw n
		}
		try {
			var i = wb[r] || r || "General";
			if (t.cellNF) e.z = i;
			if (t.cellDates && e.t == "n" && Le(i)) {
				var s = ae(e.v + (a ? 1462 : 0));
				if (s) {
					e.t = "d";
					e.v = new Date(Date.UTC(s.y, s.m - 1, s.d, s.H, s.M, s.S, s.u))
				}
			}
		} catch(n) {
			if (t.WTF) throw n
		}
	}
	function Eb(e, r, t) {
		if (t.cellStyles) {
			if (r.Interior) {
				var a = r.Interior;
				if (a.Pattern) a.patternType = oc[a.Pattern] || a.Pattern
			}
		}
		e[r.ID] = r
	}
	function _b(e, r, t, a, n, i, s, f, l, o, c) {
		var u = "General",
		h = a.StyleID,
		d = {};
		o = o || {};
		var v = [];
		var p = 0;
		if (h === undefined && f) h = f.StyleID;
		if (h === undefined && s) h = s.StyleID;
		while (i[h] !== undefined) {
			if (i[h].nf) u = i[h].nf;
			if (i[h].Interior) v.push(i[h].Interior);
			if (!i[h].Parent) break;
			h = i[h].Parent
		}
		switch (t.Type) {
		case "Boolean":
			a.t = "b";
			a.v = pt(e);
			break;
		case "String":
			a.t = "s";
			a.r = dt(it(e));
			a.v = e.indexOf("<") > -1 ? it(r || e).replace(/<.*?>/g, "") : a.r;
			break;
		case "DateTime":
			if (e.slice(-1) != "Z") e += "Z";
			a.v = dr(wr(e, c), c);
			if (a.v !== a.v) a.v = it(e);
			if (!u || u == "General") u = "yyyy-mm-dd";
		case "Number":
			if (a.v === undefined) a.v = +e;
			if (!a.t) a.t = "n";
			break;
		case "Error":
			a.t = "e";
			a.v = ai[e];
			if (o.cellText !== false) a.w = e;
			break;
		default:
			if (e == "" && r == "") {
				a.t = "z"
			} else {
				a.t = "s";
				a.v = dt(r || e)
			}
			break;
		}
		yb(a, u, o, c);
		if (o.cellFormula !== false) {
			if (a.Formula) {
				var m = it(a.Formula);
				if (m.charCodeAt(0) == 61) m = m.slice(1);
				a.f = rh(m, n);
				delete a.Formula;
				if (a.ArrayRange == "RC") a.F = rh("RC:RC", n);
				else if (a.ArrayRange) {
					a.F = rh(a.ArrayRange, n);
					l.push([Ga(a.F), a.F])
				}
			} else {
				for (p = 0; p < l.length; ++p) if (n.r >= l[p][0].s.r && n.r <= l[p][0].e.r) if (n.c >= l[p][0].s.c && n.c <= l[p][0].e.c) a.F = l[p][1]
			}
		}
		if (o.cellStyles) {
			v.forEach(function(e) {
				if (!d.patternType && e.patternType) d.patternType = e.patternType
			});
			a.s = d
		}
		if (a.StyleID !== undefined) a.ixfe = a.StyleID
	}
	function Sb(e) {
		return ni.indexOf("_xlnm." + e) > -1 ? "_xlnm." + e: e
	}
	function xb(e) {
		e.t = e.v || "";
		e.t = e.t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
		e.v = e.w = e.ixfe = undefined
	}
	function Ab(e, r) {
		var t = r || {};
		$e();
		var n = v(Dt(e));
		if (t.type == "binary" || t.type == "array" || t.type == "base64") {
			if (typeof a !== "undefined") n = a.utils.decode(65001, c(n));
			else n = kt(n)
		}
		var i = n.slice(0, 1024).toLowerCase(),
		s = false;
		i = i.replace(/".*?"/g, "");
		if ((i.indexOf(">") & 1023) > Math.min(i.indexOf(",") & 1023, i.indexOf(";") & 1023)) {
			var f = Tr(t);
			f.type = "string";
			return Yl.to_workbook(n, f)
		}
		if (i.indexOf("<?xml") == -1)["html", "table", "head", "meta", "script", "style", "div"].forEach(function(e) {
			if (i.indexOf("<" + e) >= 0) s = true
		});
		if (s) return Cg(n, t);
		wb = {
			"General Number": "General",
			"General Date": q[22],
			"Long Date": "dddd, mmmm dd, yyyy",
			"Medium Date": q[15],
			"Short Date": q[14],
			"Long Time": q[19],
			"Medium Time": q[18],
			"Short Time": q[20],
			Currency: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
			Fixed: q[2],
			Standard: q[4],
			Percent: q[10],
			Scientific: q[11],
			"Yes/No": '"Yes";"Yes";"No";@',
			"True/False": '"True";"True";"False";@',
			"On/Off": '"Yes";"Yes";"No";@'
		};
		var l;
		var o = [],
		u;
		if (g != null && t.dense == null) t.dense = g;
		var h = {},
		d = [],
		p = {},
		m = "";
		if (t.dense) p["!data"] = [];
		var b = {},
		w = {};
		var k = bb('<Data ss:Type="String">'),
		T = 0;
		var y = 0,
		E = 0;
		var _ = {
			s: {
				r: 2e6,
				c: 2e6
			},
			e: {
				r: 0,
				c: 0
			}
		};
		var S = {},
		x = {};
		var A = "",
		C = 0;
		var R = [];
		var O = {},
		I = {},
		N = 0,
		F = [];
		var D = [],
		P = {};
		var L = [],
		M,
		U = false;
		var B = [];
		var W = [],
		z = {},
		H = 0,
		V = 0;
		var $ = {
			Sheets: [],
			WBProps: {
				date1904: false
			}
		},
		X = {};
		Pt.lastIndex = 0;
		n = n.replace(/<!--([\s\S]*?)-->/gm, "");
		var G = "";
		while (l = Pt.exec(n)) switch (l[3] = (G = l[3]).toLowerCase()) {
		case "data":
			if (G == "data") {
				if (l[1] === "/") {
					if ((u = o.pop())[0] !== l[3]) throw new Error("Bad state: " + u.join("|"))
				} else if (l[0].charAt(l[0].length - 2) !== "/") o.push([l[3], true]);
				break
			}
			if (o[o.length - 1][1]) break;
			if (l[1] === "/") _b(n.slice(T, l.index), A, k, o[o.length - 1][0] == "comment" ? P: b, {
				c: y,
				r: E
			},
			S, L[y], w, B, t, $.WBProps.date1904);
			else {
				A = "";
				k = bb(l[0]);
				T = l.index + l[0].length
			}
			break;
		case "cell":
			if (l[1] === "/") {
				if (D.length > 0) b.c = D;
				if ((!t.sheetRows || t.sheetRows > E) && b.v !== void 0) {
					if (t.dense) {
						if (!p["!data"][E]) p["!data"][E] = [];
						p["!data"][E][y] = b
					} else p[La(y) + Na(E)] = b
				}
				if (b.HRef) {
					b.l = {
						Target: it(b.HRef)
					};
					if (b.HRefScreenTip) b.l.Tooltip = b.HRefScreenTip;
					delete b.HRef;
					delete b.HRefScreenTip
				}
				if (b.MergeAcross || b.MergeDown) {
					H = y + (parseInt(b.MergeAcross, 10) | 0);
					V = E + (parseInt(b.MergeDown, 10) | 0);
					if (H > y || V > E) R.push({
						s: {
							c: y,
							r: E
						},
						e: {
							c: H,
							r: V
						}
					})
				}
				if (!t.sheetStubs) {
					if (b.MergeAcross) y = H + 1;
					else++y
				} else if (b.MergeAcross || b.MergeDown) {
					for (var j = y; j <= H; ++j) {
						for (var K = E; K <= V; ++K) {
							if (j > y || K > E) {
								if (t.dense) {
									if (!p["!data"][K]) p["!data"][K] = [];
									p["!data"][K][j] = {
										t: "z"
									}
								} else p[La(j) + Na(K)] = {
									t: "z"
								}
							}
						}
					}
					y = H + 1
				} else++y
			} else {
				b = gb(l[0]);
				if (b.Index) y = +b.Index - 1;
				if (y < _.s.c) _.s.c = y;
				if (y > _.e.c) _.e.c = y;
				if (l[0].slice(-2) === "/>")++y;
				D = []
			}
			break;
		case "row":
			if (l[1] === "/" || l[0].slice(-2) === "/>") {
				if (E < _.s.r) _.s.r = E;
				if (E > _.e.r) _.e.r = E;
				if (l[0].slice(-2) === "/>") {
					w = bb(l[0]);
					if (w.Index) E = +w.Index - 1
				}
				y = 0; ++E
			} else {
				w = bb(l[0]);
				if (w.Index) E = +w.Index - 1;
				z = {};
				if (w.AutoFitHeight == "0" || w.Height) {
					z.hpx = parseInt(w.Height, 10);
					z.hpt = fc(z.hpx);
					W[E] = z
				}
				if (w.Hidden == "1") {
					z.hidden = true;
					W[E] = z
				}
			}
			break;
		case "worksheet":
			if (l[1] === "/") {
				if ((u = o.pop())[0] !== l[3]) throw new Error("Bad state: " + u.join("|"));
				d.push(m);
				if (_.s.r <= _.e.r && _.s.c <= _.e.c) {
					p["!ref"] = Va(_);
					if (t.sheetRows && t.sheetRows <= _.e.r) {
						p["!fullref"] = p["!ref"];
						_.e.r = t.sheetRows - 1;
						p["!ref"] = Va(_)
					}
				}
				if (R.length) p["!merges"] = R;
				if (L.length > 0) p["!cols"] = L;
				if (W.length > 0) p["!rows"] = W;
				h[m] = p
			} else {
				_ = {
					s: {
						r: 2e6,
						c: 2e6
					},
					e: {
						r: 0,
						c: 0
					}
				};
				E = y = 0;
				o.push([l[3], false]);
				u = bb(l[0]);
				m = it(u.Name);
				p = {};
				if (t.dense) p["!data"] = [];
				R = [];
				B = [];
				W = [];
				X = {
					name: m,
					Hidden: 0
				};
				$.Sheets.push(X)
			}
			break;
		case "table":
			if (l[1] === "/") {
				if ((u = o.pop())[0] !== l[3]) throw new Error("Bad state: " + u.join("|"))
			} else if (l[0].slice(-2) == "/>") break;
			else {
				o.push([l[3], false]);
				L = [];
				U = false
			}
			break;
		case "style":
			if (l[1] === "/") Eb(S, x, t);
			else x = bb(l[0]);
			break;
		case "numberformat":
			x.nf = it(bb(l[0]).Format || "General");
			if (wb[x.nf]) x.nf = wb[x.nf];
			for (var Y = 0; Y != 392; ++Y) if (q[Y] == x.nf) break;
			if (Y == 392) for (Y = 57; Y != 392; ++Y) if (q[Y] == null) {
				Je(x.nf, Y);
				break
			}
			break;
		case "column":
			if (o[o.length - 1][0] !== "table") break;
			if (l[1] === "/") break;
			M = bb(l[0]);
			if (M.Hidden) {
				M.hidden = true;
				delete M.Hidden
			}
			if (M.Width) M.wpx = parseInt(M.Width, 10);
			if (!U && M.wpx > 10) {
				U = true;
				qo = Yo;
				for (var Z = 0; Z < L.length; ++Z) if (L[Z]) nc(L[Z])
			}
			if (U) nc(M);
			L[M.Index - 1 || L.length] = M;
			for (var J = 0; J < +M.Span; ++J) L[L.length] = Tr(M);
			break;
		case "namedrange":
			if (l[1] === "/") break;
			if (!$.Names) $.Names = [];
			var Q = rt(l[0]);
			var ee = {
				Name: Sb(Q.Name),
				Ref: rh(Q.RefersTo.slice(1), {
					r: 0,
					c: 0
				})
			};
			if ($.Sheets.length > 0) ee.Sheet = $.Sheets.length - 1;
			$.Names.push(ee);
			break;
		case "namedcell":
			break;
		case "b":
			break;
		case "i":
			break;
		case "u":
			break;
		case "s":
			break;
		case "em":
			break;
		case "h2":
			break;
		case "h3":
			break;
		case "sub":
			break;
		case "sup":
			break;
		case "span":
			break;
		case "alignment":
			break;
		case "borders":
			break;
		case "border":
			break;
		case "font":
			if (l[0].slice(-2) === "/>") break;
			else if (l[1] === "/") A += n.slice(C, l.index);
			else C = l.index + l[0].length;
			break;
		case "interior":
			if (!t.cellStyles) break;
			x.Interior = bb(l[0]);
			break;
		case "protection":
			break;
		case "author":
			;
		case "title":
			;
		case "description":
			;
		case "created":
			;
		case "keywords":
			;
		case "subject":
			;
		case "category":
			;
		case "company":
			;
		case "lastauthor":
			;
		case "lastsaved":
			;
		case "lastprinted":
			;
		case "version":
			;
		case "revision":
			;
		case "totaltime":
			;
		case "hyperlinkbase":
			;
		case "manager":
			;
		case "contentstatus":
			;
		case "identifier":
			;
		case "language":
			;
		case "appname":
			if (l[0].slice(-2) === "/>") break;
			else if (l[1] === "/") Mi(O, G, n.slice(N, l.index));
			else N = l.index + l[0].length;
			break;
		case "paragraphs":
			break;
		case "styles":
			;
		case "workbook":
			if (l[1] === "/") {
				if ((u = o.pop())[0] !== l[3]) throw new Error("Bad state: " + u.join("|"))
			} else o.push([l[3], false]);
			break;
		case "comment":
			if (l[1] === "/") {
				if ((u = o.pop())[0] !== l[3]) throw new Error("Bad state: " + u.join("|"));
				xb(P);
				D.push(P)
			} else {
				o.push([l[3], false]);
				u = bb(l[0]);
				if (!pt(u["ShowAlways"] || "0")) D.hidden = true;
				P = {
					a: u.Author
				}
			}
			break;
		case "autofilter":
			if (l[1] === "/") {
				if ((u = o.pop())[0] !== l[3]) throw new Error("Bad state: " + u.join("|"))
			} else if (l[0].charAt(l[0].length - 2) !== "/") {
				var re = bb(l[0]);
				p["!autofilter"] = {
					ref: rh(re.Range).replace(/\$/g, "")
				};
				o.push([l[3], true])
			}
			break;
		case "name":
			break;
		case "datavalidation":
			if (l[1] === "/") {
				if ((u = o.pop())[0] !== l[3]) throw new Error("Bad state: " + u.join("|"))
			} else {
				if (l[0].charAt(l[0].length - 2) !== "/") o.push([l[3], true])
			}
			break;
		case "pixelsperinch":
			break;
		case "componentoptions":
			;
		case "documentproperties":
			;
		case "customdocumentproperties":
			;
		case "officedocumentsettings":
			;
		case "pivottable":
			;
		case "pivotcache":
			;
		case "names":
			;
		case "mapinfo":
			;
		case "pagebreaks":
			;
		case "querytable":
			;
		case "sorting":
			;
		case "schema":
			;
		case "conditionalformatting":
			;
		case "smarttagtype":
			;
		case "smarttags":
			;
		case "excelworkbook":
			;
		case "workbookoptions":
			;
		case "worksheetoptions":
			if (l[1] === "/") {
				if ((u = o.pop())[0] !== l[3]) throw new Error("Bad state: " + u.join("|"))
			} else if (l[0].charAt(l[0].length - 2) !== "/") o.push([l[3], true]);
			break;
		case "null":
			break;
		default:
			if (o.length == 0 && l[3] == "document") return Wg(n, t);
			if (o.length == 0 && l[3] == "uof") return Wg(n, t);
			var te = true;
			switch (o[o.length - 1][0]) {
			case "officedocumentsettings":
				switch (l[3]) {
				case "allowpng":
					break;
				case "removepersonalinformation":
					break;
				case "downloadcomponents":
					break;
				case "locationofcomponents":
					break;
				case "colors":
					break;
				case "color":
					break;
				case "index":
					break;
				case "rgb":
					break;
				case "targetscreensize":
					break;
				case "readonlyrecommended":
					break;
				default:
					te = false;
				}
				break;
			case "componentoptions":
				switch (l[3]) {
				case "toolbar":
					break;
				case "hideofficelogo":
					break;
				case "spreadsheetautofit":
					break;
				case "label":
					break;
				case "caption":
					break;
				case "maxheight":
					break;
				case "maxwidth":
					break;
				case "nextsheetnumber":
					break;
				default:
					te = false;
				}
				break;
			case "excelworkbook":
				switch (l[3]) {
				case "date1904":
					$.WBProps.date1904 = true;
					break;
				case "hidehorizontalscrollbar":
					break;
				case "hideverticalscrollbar":
					break;
				case "hideworkbooktabs":
					break;
				case "windowheight":
					break;
				case "windowwidth":
					break;
				case "windowtopx":
					break;
				case "windowtopy":
					break;
				case "tabratio":
					break;
				case "protectstructure":
					break;
				case "protectwindow":
					break;
				case "protectwindows":
					break;
				case "activesheet":
					break;
				case "displayinknotes":
					break;
				case "firstvisiblesheet":
					break;
				case "supbook":
					break;
				case "sheetname":
					break;
				case "sheetindex":
					break;
				case "sheetindexfirst":
					break;
				case "sheetindexlast":
					break;
				case "dll":
					break;
				case "acceptlabelsinformulas":
					break;
				case "donotsavelinkvalues":
					break;
				case "iteration":
					break;
				case "maxiterations":
					break;
				case "maxchange":
					break;
				case "path":
					break;
				case "xct":
					break;
				case "count":
					break;
				case "selectedsheets":
					break;
				case "calculation":
					break;
				case "uncalced":
					break;
				case "startupprompt":
					break;
				case "crn":
					break;
				case "externname":
					break;
				case "formula":
					break;
				case "colfirst":
					break;
				case "collast":
					break;
				case "wantadvise":
					break;
				case "boolean":
					break;
				case "error":
					break;
				case "text":
					break;
				case "ole":
					break;
				case "noautorecover":
					break;
				case "publishobjects":
					break;
				case "donotcalculatebeforesave":
					break;
				case "number":
					break;
				case "refmoder1c1":
					break;
				case "embedsavesmarttags":
					break;
				default:
					te = false;
				}
				break;
			case "workbookoptions":
				switch (l[3]) {
				case "owcversion":
					break;
				case "height":
					break;
				case "width":
					break;
				default:
					te = false;
				}
				break;
			case "worksheetoptions":
				switch (l[3]) {
				case "visible":
					if (l[0].slice(-2) === "/>") {} else if (l[1] === "/") switch (n.slice(N, l.index)) {
					case "SheetHidden":
						X.Hidden = 1;
						break;
					case "SheetVeryHidden":
						X.Hidden = 2;
						break;
					} else N = l.index + l[0].length;
					break;
				case "header":
					if (!p["!margins"]) Tv(p["!margins"] = {},
					"xlml");
					if (!isNaN( + rt(l[0]).Margin)) p["!margins"].header = +rt(l[0]).Margin;
					break;
				case "footer":
					if (!p["!margins"]) Tv(p["!margins"] = {},
					"xlml");
					if (!isNaN( + rt(l[0]).Margin)) p["!margins"].footer = +rt(l[0]).Margin;
					break;
				case "pagemargins":
					var ae = rt(l[0]);
					if (!p["!margins"]) Tv(p["!margins"] = {},
					"xlml");
					if (!isNaN( + ae.Top)) p["!margins"].top = +ae.Top;
					if (!isNaN( + ae.Left)) p["!margins"].left = +ae.Left;
					if (!isNaN( + ae.Right)) p["!margins"].right = +ae.Right;
					if (!isNaN( + ae.Bottom)) p["!margins"].bottom = +ae.Bottom;
					break;
				case "displayrighttoleft":
					if (!$.Views) $.Views = [];
					if (!$.Views[0]) $.Views[0] = {};
					$.Views[0].RTL = true;
					break;
				case "freezepanes":
					break;
				case "frozennosplit":
					break;
				case "splithorizontal":
					;
				case "splitvertical":
					break;
				case "donotdisplaygridlines":
					break;
				case "activerow":
					break;
				case "activecol":
					break;
				case "toprowbottompane":
					break;
				case "leftcolumnrightpane":
					break;
				case "unsynced":
					break;
				case "print":
					break;
				case "printerrors":
					break;
				case "panes":
					break;
				case "scale":
					break;
				case "pane":
					break;
				case "number":
					break;
				case "layout":
					break;
				case "pagesetup":
					break;
				case "selected":
					break;
				case "protectobjects":
					break;
				case "enableselection":
					break;
				case "protectscenarios":
					break;
				case "validprinterinfo":
					break;
				case "horizontalresolution":
					break;
				case "verticalresolution":
					break;
				case "numberofcopies":
					break;
				case "activepane":
					break;
				case "toprowvisible":
					break;
				case "leftcolumnvisible":
					break;
				case "fittopage":
					break;
				case "rangeselection":
					break;
				case "papersizeindex":
					break;
				case "pagelayoutzoom":
					break;
				case "pagebreakzoom":
					break;
				case "filteron":
					break;
				case "fitwidth":
					break;
				case "fitheight":
					break;
				case "commentslayout":
					break;
				case "zoom":
					break;
				case "lefttoright":
					break;
				case "gridlines":
					break;
				case "allowsort":
					break;
				case "allowfilter":
					break;
				case "allowinsertrows":
					break;
				case "allowdeleterows":
					break;
				case "allowinsertcols":
					break;
				case "allowdeletecols":
					break;
				case "allowinserthyperlinks":
					break;
				case "allowformatcells":
					break;
				case "allowsizecols":
					break;
				case "allowsizerows":
					break;
				case "nosummaryrowsbelowdetail":
					if (!p["!outline"]) p["!outline"] = {};
					p["!outline"].above = true;
					break;
				case "tabcolorindex":
					break;
				case "donotdisplayheadings":
					break;
				case "showpagelayoutzoom":
					break;
				case "nosummarycolumnsrightdetail":
					if (!p["!outline"]) p["!outline"] = {};
					p["!outline"].left = true;
					break;
				case "blackandwhite":
					break;
				case "donotdisplayzeros":
					break;
				case "displaypagebreak":
					break;
				case "rowcolheadings":
					break;
				case "donotdisplayoutline":
					break;
				case "noorientation":
					break;
				case "allowusepivottables":
					break;
				case "zeroheight":
					break;
				case "viewablerange":
					break;
				case "selection":
					break;
				case "protectcontents":
					break;
				default:
					te = false;
				}
				break;
			case "pivottable":
				;
			case "pivotcache":
				switch (l[3]) {
				case "immediateitemsondrop":
					break;
				case "showpagemultipleitemlabel":
					break;
				case "compactrowindent":
					break;
				case "location":
					break;
				case "pivotfield":
					break;
				case "orientation":
					break;
				case "layoutform":
					break;
				case "layoutsubtotallocation":
					break;
				case "layoutcompactrow":
					break;
				case "position":
					break;
				case "pivotitem":
					break;
				case "datatype":
					break;
				case "datafield":
					break;
				case "sourcename":
					break;
				case "parentfield":
					break;
				case "ptlineitems":
					break;
				case "ptlineitem":
					break;
				case "countofsameitems":
					break;
				case "item":
					break;
				case "itemtype":
					break;
				case "ptsource":
					break;
				case "cacheindex":
					break;
				case "consolidationreference":
					break;
				case "filename":
					break;
				case "reference":
					break;
				case "nocolumngrand":
					break;
				case "norowgrand":
					break;
				case "blanklineafteritems":
					break;
				case "hidden":
					break;
				case "subtotal":
					break;
				case "basefield":
					break;
				case "mapchilditems":
					break;
				case "function":
					break;
				case "refreshonfileopen":
					break;
				case "printsettitles":
					break;
				case "mergelabels":
					break;
				case "defaultversion":
					break;
				case "refreshname":
					break;
				case "refreshdate":
					break;
				case "refreshdatecopy":
					break;
				case "versionlastrefresh":
					break;
				case "versionlastupdate":
					break;
				case "versionupdateablemin":
					break;
				case "versionrefreshablemin":
					break;
				case "calculation":
					break;
				default:
					te = false;
				}
				break;
			case "pagebreaks":
				switch (l[3]) {
				case "colbreaks":
					break;
				case "colbreak":
					break;
				case "rowbreaks":
					break;
				case "rowbreak":
					break;
				case "colstart":
					break;
				case "colend":
					break;
				case "rowend":
					break;
				default:
					te = false;
				}
				break;
			case "autofilter":
				switch (l[3]) {
				case "autofiltercolumn":
					break;
				case "autofiltercondition":
					break;
				case "autofilterand":
					break;
				case "autofilteror":
					break;
				default:
					te = false;
				}
				break;
			case "querytable":
				switch (l[3]) {
				case "id":
					break;
				case "autoformatfont":
					break;
				case "autoformatpattern":
					break;
				case "querysource":
					break;
				case "querytype":
					break;
				case "enableredirections":
					break;
				case "refreshedinxl9":
					break;
				case "urlstring":
					break;
				case "htmltables":
					break;
				case "connection":
					break;
				case "commandtext":
					break;
				case "refreshinfo":
					break;
				case "notitles":
					break;
				case "nextid":
					break;
				case "columninfo":
					break;
				case "overwritecells":
					break;
				case "donotpromptforfile":
					break;
				case "textwizardsettings":
					break;
				case "source":
					break;
				case "number":
					break;
				case "decimal":
					break;
				case "thousandseparator":
					break;
				case "trailingminusnumbers":
					break;
				case "formatsettings":
					break;
				case "fieldtype":
					break;
				case "delimiters":
					break;
				case "tab":
					break;
				case "comma":
					break;
				case "autoformatname":
					break;
				case "versionlastedit":
					break;
				case "versionlastrefresh":
					break;
				default:
					te = false;
				}
				break;
			case "datavalidation":
				switch (l[3]) {
				case "range":
					break;
				case "type":
					break;
				case "min":
					break;
				case "max":
					break;
				case "sort":
					break;
				case "descending":
					break;
				case "order":
					break;
				case "casesensitive":
					break;
				case "value":
					break;
				case "errorstyle":
					break;
				case "errormessage":
					break;
				case "errortitle":
					break;
				case "inputmessage":
					break;
				case "inputtitle":
					break;
				case "combohide":
					break;
				case "inputhide":
					break;
				case "condition":
					break;
				case "qualifier":
					break;
				case "useblank":
					break;
				case "value1":
					break;
				case "value2":
					break;
				case "format":
					break;
				case "cellrangelist":
					break;
				default:
					te = false;
				}
				break;
			case "sorting":
				;
			case "conditionalformatting":
				switch (l[3]) {
				case "range":
					break;
				case "type":
					break;
				case "min":
					break;
				case "max":
					break;
				case "sort":
					break;
				case "descending":
					break;
				case "order":
					break;
				case "casesensitive":
					break;
				case "value":
					break;
				case "errorstyle":
					break;
				case "errormessage":
					break;
				case "errortitle":
					break;
				case "cellrangelist":
					break;
				case "inputmessage":
					break;
				case "inputtitle":
					break;
				case "combohide":
					break;
				case "inputhide":
					break;
				case "condition":
					break;
				case "qualifier":
					break;
				case "useblank":
					break;
				case "value1":
					break;
				case "value2":
					break;
				case "format":
					break;
				default:
					te = false;
				}
				break;
			case "mapinfo":
				;
			case "schema":
				;
			case "data":
				switch (l[3]) {
				case "map":
					break;
				case "entry":
					break;
				case "range":
					break;
				case "xpath":
					break;
				case "field":
					break;
				case "xsdtype":
					break;
				case "filteron":
					break;
				case "aggregate":
					break;
				case "elementtype":
					break;
				case "attributetype":
					break;
				case "schema":
					;
				case "element":
					;
				case "complextype":
					;
				case "datatype":
					;
				case "all":
					;
				case "attribute":
					;
				case "extends":
					break;
				case "row":
					break;
				default:
					te = false;
				}
				break;
			case "smarttags":
				break;
			default:
				te = false;
				break;
			}
			if (te) break;
			if (l[3].match(/!\[CDATA/)) break;
			if (!o[o.length - 1][1]) throw "Unrecognized tag: " + l[3] + "|" + o.join("|");
			if (o[o.length - 1][0] === "customdocumentproperties") {
				if (l[0].slice(-2) === "/>") break;
				else if (l[1] === "/") Tb(I, G, F, n.slice(N, l.index));
				else {
					F = l;
					N = l.index + l[0].length
				}
				break
			}
			if (t.WTF) throw "Unrecognized tag: " + l[3] + "|" + o.join("|");
		}
		var ne = {};
		if (!t.bookSheets && !t.bookProps) ne.Sheets = h;
		ne.SheetNames = d;
		ne.Workbook = $;
		ne.SSF = Tr(q);
		ne.Props = O;
		ne.Custprops = I;
		ne.bookType = "xlml";
		return ne
	}
	function Cb(e, r) {
		ek(r = r || {});
		switch (r.type || "base64") {
		case "base64":
			return Ab(_(e), r);
		case "binary":
			;
		case "buffer":
			;
		case "file":
			return Ab(e, r);
		case "array":
			return Ab(N(e), r);
		}
	}
	function Rb(e, r) {
		var t = [];
		if (e.Props) t.push(Ui(e.Props, r));
		if (e.Custprops) t.push(Bi(e.Props, e.Custprops, r));
		return t.join("")
	}
	function Ob(e) {
		if ((((e || {}).Workbook || {}).WBProps || {}).date1904) return '<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel"><Date1904/></ExcelWorkbook>';
		return ""
	}
	function Ib(e, r) {
		var t = ['<Style ss:ID="Default" ss:Name="Normal"><NumberFormat/></Style>'];
		r.cellXfs.forEach(function(e, r) {
			var a = [];
			a.push(It("NumberFormat", null, {
				"ss:Format": lt(q[e.numFmtId])
			}));
			var n = {
				"ss:ID": "s" + (21 + r)
			};
			t.push(It("Style", a.join(""), n))
		});
		return It("Styles", t.join(""))
	}
	function Nb(e) {
		return It("NamedRange", null, {
			"ss:Name": e.Name.slice(0, 6) == "_xlnm." ? e.Name.slice(6) : e.Name,
			"ss:RefersTo": "=" + nh(e.Ref, {
				r: 0,
				c: 0
			})
		})
	}
	function Fb(e) {
		if (! ((e || {}).Workbook || {}).Names) return "";
		var r = e.Workbook.Names;
		var t = [];
		for (var a = 0; a < r.length; ++a) {
			var n = r[a];
			if (n.Sheet != null) continue;
			if (n.Name.match(/^_xlfn\./)) continue;
			t.push(Nb(n))
		}
		return It("Names", t.join(""))
	}
	function Db(e, r, t, a) {
		if (!e) return "";
		if (! ((a || {}).Workbook || {}).Names) return "";
		var n = a.Workbook.Names;
		var i = [];
		for (var s = 0; s < n.length; ++s) {
			var f = n[s];
			if (f.Sheet != t) continue;
			if (f.Name.match(/^_xlfn\./)) continue;
			i.push(Nb(f))
		}
		return i.join("")
	}
	function Pb(e, r, t, a) {
		if (!e) return "";
		var n = [];
		if (e["!margins"]) {
			n.push("<PageSetup>");
			if (e["!margins"].header) n.push(It("Header", null, {
				"x:Margin": e["!margins"].header
			}));
			if (e["!margins"].footer) n.push(It("Footer", null, {
				"x:Margin": e["!margins"].footer
			}));
			n.push(It("PageMargins", null, {
				"x:Bottom": e["!margins"].bottom || "0.75",
				"x:Left": e["!margins"].left || "0.7",
				"x:Right": e["!margins"].right || "0.7",
				"x:Top": e["!margins"].top || "0.75"
			}));
			n.push("</PageSetup>")
		}
		if (a && a.Workbook && a.Workbook.Sheets && a.Workbook.Sheets[t]) {
			if (a.Workbook.Sheets[t].Hidden) n.push(It("Visible", a.Workbook.Sheets[t].Hidden == 1 ? "SheetHidden": "SheetVeryHidden", {}));
			else {
				for (var i = 0; i < t; ++i) if (a.Workbook.Sheets[i] && !a.Workbook.Sheets[i].Hidden) break;
				if (i == t) n.push("<Selected/>")
			}
		}
		if (((((a || {}).Workbook || {}).Views || [])[0] || {}).RTL) n.push("<DisplayRightToLeft/>");
		if (e["!protect"]) {
			n.push(Rt("ProtectContents", "True"));
			if (e["!protect"].objects) n.push(Rt("ProtectObjects", "True"));
			if (e["!protect"].scenarios) n.push(Rt("ProtectScenarios", "True"));
			if (e["!protect"].selectLockedCells != null && !e["!protect"].selectLockedCells) n.push(Rt("EnableSelection", "NoSelection"));
			else if (e["!protect"].selectUnlockedCells != null && !e["!protect"].selectUnlockedCells) n.push(Rt("EnableSelection", "UnlockedCells")); [["formatCells", "AllowFormatCells"], ["formatColumns", "AllowSizeCols"], ["formatRows", "AllowSizeRows"], ["insertColumns", "AllowInsertCols"], ["insertRows", "AllowInsertRows"], ["insertHyperlinks", "AllowInsertHyperlinks"], ["deleteColumns", "AllowDeleteCols"], ["deleteRows", "AllowDeleteRows"], ["sort", "AllowSort"], ["autoFilter", "AllowFilter"], ["pivotTables", "AllowUsePivotTables"]].forEach(function(r) {
				if (e["!protect"][r[0]]) n.push("<" + r[1] + "/>")
			})
		}
		if (n.length == 0) return "";
		return It("WorksheetOptions", n.join(""), {
			xmlns: Ut.x
		})
	}
	function Lb(e) {
		return e.map(function(r) {
			var t = vt(r.t || "");
			var a = It("ss:Data", t, {
				xmlns: "http://www.w3.org/TR/REC-html40"
			});
			var n = {};
			if (r.a) n["ss:Author"] = r.a;
			if (!e.hidden) n["ss:ShowAlways"] = "1";
			return It("Comment", a, n)
		}).join("")
	}
	function Mb(e, r, t, a, n, i, s) {
		if (!e || e.v == undefined && e.f == undefined) return "";
		var f = {};
		if (e.f) f["ss:Formula"] = "=" + lt(nh(e.f, s));
		if (e.F && e.F.slice(0, r.length) == r) {
			var l = Wa(e.F.slice(r.length + 1));
			f["ss:ArrayRange"] = "RC:R" + (l.r == s.r ? "": "[" + (l.r - s.r) + "]") + "C" + (l.c == s.c ? "": "[" + (l.c - s.c) + "]")
		}
		if (e.l && e.l.Target) {
			f["ss:HRef"] = lt(e.l.Target);
			if (e.l.Tooltip) f["x:HRefScreenTip"] = lt(e.l.Tooltip)
		}
		if (t["!merges"]) {
			var o = t["!merges"];
			for (var c = 0; c != o.length; ++c) {
				if (o[c].s.c != s.c || o[c].s.r != s.r) continue;
				if (o[c].e.c > o[c].s.c) f["ss:MergeAcross"] = o[c].e.c - o[c].s.c;
				if (o[c].e.r > o[c].s.r) f["ss:MergeDown"] = o[c].e.r - o[c].s.r
			}
		}
		var u = "",
		h = "";
		switch (e.t) {
		case "z":
			if (!a.sheetStubs) return "";
			break;
		case "n":
			u = "Number";
			h = String(e.v);
			break;
		case "b":
			u = "Boolean";
			h = e.v ? "1": "0";
			break;
		case "e":
			u = "Error";
			h = ti[e.v];
			break;
		case "d":
			u = "DateTime";
			h = new Date(e.v).toISOString();
			if (e.z == null) e.z = e.z || q[14];
			break;
		case "s":
			u = "String";
			h = ht(e.v || "");
			break;
		}
		var d = yv(a.cellXfs, e, a);
		f["ss:StyleID"] = "s" + (21 + d);
		f["ss:Index"] = s.c + 1;
		var v = e.v != null ? h: "";
		var p = e.t == "z" ? "": '<Data ss:Type="' + u + '">' + v + "</Data>";
		if ((e.c || []).length > 0) p += Lb(e.c);
		return It("Cell", p, f)
	}
	function Ub(e, r) {
		var t = '<Row ss:Index="' + (e + 1) + '"';
		if (r) {
			if (r.hpt && !r.hpx) r.hpx = lc(r.hpt);
			if (r.hpx) t += ' ss:AutoFitHeight="0" ss:Height="' + r.hpx + '"';
			if (r.hidden) t += ' ss:Hidden="1"'
		}
		return t + ">"
	}
	function Bb(e, r, t, a) {
		if (!e["!ref"]) return "";
		var n = Ga(e["!ref"]);
		var i = e["!merges"] || [],
		s = 0;
		var f = [];
		if (e["!cols"]) e["!cols"].forEach(function(e, r) {
			nc(e);
			var t = !!e.width;
			var a = kv(r, e);
			var n = {
				"ss:Index": r + 1
			};
			if (t) n["ss:Width"] = Qo(a.width);
			if (e.hidden) n["ss:Hidden"] = "1";
			f.push(It("Column", null, n))
		});
		var l = e["!data"] != null;
		for (var o = n.s.r; o <= n.e.r; ++o) {
			var c = [Ub(o, (e["!rows"] || [])[o])];
			for (var u = n.s.c; u <= n.e.c; ++u) {
				var h = false;
				for (s = 0; s != i.length; ++s) {
					if (i[s].s.c > u) continue;
					if (i[s].s.r > o) continue;
					if (i[s].e.c < u) continue;
					if (i[s].e.r < o) continue;
					if (i[s].s.c != u || i[s].s.r != o) h = true;
					break
				}
				if (h) continue;
				var d = {
					r: o,
					c: u
				};
				var v = La(u) + Na(o),
				p = l ? (e["!data"][o] || [])[u] : e[v];
				c.push(Mb(p, v, e, r, t, a, d))
			}
			c.push("</Row>");
			if (c.length > 2) f.push(c.join(""))
		}
		return f.join("")
	}
	function Wb(e, r, t) {
		var a = [];
		var n = t.SheetNames[e];
		var i = t.Sheets[n];
		var s = i ? Db(i, r, e, t) : "";
		if (s.length > 0) a.push("<Names>" + s + "</Names>");
		s = i ? Bb(i, r, e, t) : "";
		if (s.length > 0) a.push("<Table>" + s + "</Table>");
		a.push(Pb(i, r, e, t));
		if (i["!autofilter"]) a.push('<AutoFilter x:Range="' + nh($a(i["!autofilter"].ref), {
			r: 0,
			c: 0
		}) + '" xmlns="urn:schemas-microsoft-com:office:excel"></AutoFilter>');
		return a.join("")
	}
	function zb(e, r) {
		if (!r) r = {};
		if (!e.SSF) e.SSF = Tr(q);
		if (e.SSF) {
			$e();
			Ve(e.SSF);
			r.revssf = lr(e.SSF);
			r.revssf[e.SSF[65535]] = 0;
			r.ssf = e.SSF;
			r.cellXfs = [];
			yv(r.cellXfs, {},
			{
				revssf: {
					General: 0
				}
			})
		}
		var t = [];
		t.push(Rb(e, r));
		t.push(Ob(e, r));
		t.push("");
		t.push("");
		for (var a = 0; a < e.SheetNames.length; ++a) t.push(It("Worksheet", Wb(a, r, e), {
			"ss:Name": lt(e.SheetNames[a])
		}));
		t[2] = Ib(e, r);
		t[3] = Fb(e, r);
		return Kr + It("Workbook", t.join(""), {
			xmlns: Ut.ss,
			"xmlns:o": Ut.o,
			"xmlns:x": Ut.x,
			"xmlns:ss": Ut.ss,
			"xmlns:dt": Ut.dt,
			"xmlns:html": Ut.html
		})
	}
	function Hb(e) {
		var r = {};
		var t = e.content;
		t.l = 28;
		r.AnsiUserType = t._R(0, "lpstr-ansi");
		r.AnsiClipboardFormat = Dn(t);
		if (t.length - t.l <= 4) return r;
		var a = t._R(4);
		if (a == 0 || a > 40) return r;
		t.l -= 4;
		r.Reserved1 = t._R(0, "lpstr-ansi");
		if (t.length - t.l <= 4) return r;
		a = t._R(4);
		if (a !== 1907505652) return r;
		r.UnicodeClipboardFormat = Pn(t);
		a = t._R(4);
		if (a == 0 || a > 40) return r;
		t.l -= 4;
		r.Reserved2 = t._R(0, "lpwstr")
	}
	var Vb = [60, 1084, 2066, 2165, 2175];
	function $b(e, r, t, a, n) {
		var i = a;
		var s = [];
		var f = t.slice(t.l, t.l + i);
		if (n && n.enc && n.enc.insitu && f.length > 0) switch (e) {
		case 9:
			;
		case 521:
			;
		case 1033:
			;
		case 2057:
			;
		case 47:
			;
		case 405:
			;
		case 225:
			;
		case 406:
			;
		case 312:
			;
		case 404:
			;
		case 10:
			break;
		case 133:
			break;
		default:
			n.enc.insitu(f);
		}
		s.push(f);
		t.l += i;
		var l = ca(t, t.l),
		o = eg[l];
		var c = 0;
		while (o != null && Vb.indexOf(l) > -1) {
			i = ca(t, t.l + 2);
			c = t.l + 4;
			if (l == 2066) c += 4;
			else if (l == 2165 || l == 2175) {
				c += 12
			}
			f = t.slice(c, t.l + 4 + i);
			s.push(f);
			t.l += 4 + i;
			o = eg[l = ca(t, t.l)]
		}
		var u = P(s);
		Ta(u, 0);
		var h = 0;
		u.lens = [];
		for (var d = 0; d < s.length; ++d) {
			u.lens.push(h);
			h += s[d].length
		}
		if (u.length < a) throw "XLS Record 0x" + e.toString(16) + " Truncated: " + u.length + " < " + a;
		return r.f(u, u.length, n)
	}
	function Xb(e, r, t) {
		if (e.t === "z") return;
		if (!e.XF) return;
		var a = 0;
		try {
			a = e.z || e.XF.numFmtId || 0;
			if (r.cellNF && e.z == null) e.z = q[a]
		} catch(n) {
			if (r.WTF) throw n
		}
		if (!r || r.cellText !== false) try {
			if (e.t === "e") {
				e.w = e.w || ti[e.v]
			} else if (a === 0 || a == "General") {
				if (e.t === "n") {
					if ((e.v | 0) === e.v) e.w = e.v.toString(10);
					else e.w = le(e.v)
				} else e.w = oe(e.v)
			} else e.w = ze(a, e.v, {
				date1904: !!t,
				dateNF: r && r.dateNF
			})
		} catch(n) {
			if (r.WTF) throw n
		}
		if (r.cellDates && a && e.t == "n" && Le(q[a] || String(a))) {
			var i = ae(e.v + (t ? 1462 : 0));
			if (i) {
				e.t = "d";
				e.v = new Date(Date.UTC(i.y, i.m - 1, i.d, i.H, i.M, i.S, i.u))
			}
		}
	}
	function Gb(e, r, t) {
		return {
			v: e,
			ixfe: r,
			t: t
		}
	}
	function jb(e, r) {
		var t = {
			opts: {}
		};
		var a = {};
		if (g != null && r.dense == null) r.dense = g;
		var n = {};
		if (r.dense) n["!data"] = [];
		var i = {};
		var s = {};
		var f = null;
		var o = [];
		var c = "";
		var u = {};
		var h, d = "",
		v, p, m, b;
		var w = {};
		var k = [];
		var T;
		var y;
		var E = [];
		var _ = [];
		var S = {
			Sheets: [],
			WBProps: {
				date1904: false
			},
			Views: [{}]
		},
		x = {};
		var A = false;
		var C = function pe(e) {
			if (e < 8) return ri[e];
			if (e < 64) return _[e - 8] || ri[e];
			return ri[e]
		};
		var R = function me(e, r, t) {
			var a = r.XF.data;
			if (!a || !a.patternType || !t || !t.cellStyles) return;
			r.s = {};
			r.s.patternType = a.patternType;
			var n;
			if (n = Xo(C(a.icvFore))) {
				r.s.fgColor = {
					rgb: n
				}
			}
			if (n = Xo(C(a.icvBack))) {
				r.s.bgColor = {
					rgb: n
				}
			}
		};
		var O = function be(e, r, t) {
			if (!A && W > 1) return;
			if (t.sheetRows && e.r >= t.sheetRows) return;
			if (t.cellStyles && r.XF && r.XF.data) R(e, r, t);
			delete r.ixfe;
			delete r.XF;
			h = e;
			d = za(e);
			if (!s || !s.s || !s.e) s = {
				s: {
					r: 0,
					c: 0
				},
				e: {
					r: 0,
					c: 0
				}
			};
			if (e.r < s.s.r) s.s.r = e.r;
			if (e.c < s.s.c) s.s.c = e.c;
			if (e.r + 1 > s.e.r) s.e.r = e.r + 1;
			if (e.c + 1 > s.e.c) s.e.c = e.c + 1;
			if (t.cellFormula && r.f) {
				for (var a = 0; a < k.length; ++a) {
					if (k[a][0].s.c > e.c || k[a][0].s.r > e.r) continue;
					if (k[a][0].e.c < e.c || k[a][0].e.r < e.r) continue;
					r.F = Va(k[a][0]);
					if (k[a][0].s.c != e.c || k[a][0].s.r != e.r) delete r.f;
					if (r.f) r.f = "" + Md(k[a][1], s, e, U, I);
					break
				}
			} {
				if (t.dense) {
					if (!n["!data"][e.r]) n["!data"][e.r] = [];
					n["!data"][e.r][e.c] = r
				} else n[d] = r
			}
		};
		var I = {
			enc: false,
			sbcch: 0,
			snames: [],
			sharedf: w,
			arrayf: k,
			rrtabid: [],
			lastuser: "",
			biff: 8,
			codepage: 0,
			winlocked: 0,
			cellStyles: !!r && !!r.cellStyles,
			WTF: !!r && !!r.wtf
		};
		if (r.password) I.password = r.password;
		var N;
		var F = [];
		var D = [];
		var P = [],
		L = [];
		var M = false;
		var U = [];
		U.SheetNames = I.snames;
		U.sharedf = I.sharedf;
		U.arrayf = I.arrayf;
		U.names = [];
		U.XTI = [];
		var B = 0;
		var W = 0;
		var z = 0,
		H = [];
		var V = [];
		var $;
		I.codepage = 1200;
		l(1200);
		var X = false;
		while (e.l < e.length - 1) {
			var G = e.l;
			var j = e._R(2);
			if (j === 0 && B === 10) break;
			var K = e.l === e.length ? 0 : e._R(2);
			var Y = eg[j];
			if (W == 0 && [9, 521, 1033, 2057].indexOf(j) == -1) break;
			if (Y && Y.f) {
				if (r.bookSheets) {
					if (B === 133 && j !== 133) break
				}
				B = j;
				if (Y.r === 2 || Y.r == 12) {
					var Z = e._R(2);
					K -= 2;
					if (!I.enc && Z !== j && ((Z & 255) << 8 | Z >> 8) !== j) throw new Error("rt mismatch: " + Z + "!=" + j);
					if (Y.r == 12) {
						e.l += 10;
						K -= 10
					}
				}
				var J = {};
				if (j === 10) J = Y.f(e, K, I);
				else J = $b(j, Y, e, K, I);
				if (W == 0 && [9, 521, 1033, 2057].indexOf(B) === -1) continue;
				switch (j) {
				case 34:
					t.opts.Date1904 = S.WBProps.date1904 = J;
					break;
				case 134:
					t.opts.WriteProtect = true;
					break;
				case 47:
					if (!I.enc) e.l = 0;
					I.enc = J;
					if (!r.password) throw new Error("File is password-protected");
					if (J.valid == null) throw new Error("Encryption scheme unsupported");
					if (!J.valid) throw new Error("Password is incorrect");
					break;
				case 92:
					I.lastuser = J;
					break;
				case 66:
					var Q = Number(J);
					switch (Q) {
					case 21010:
						Q = 1200;
						break;
					case 32768:
						Q = 1e4;
						break;
					case 32769:
						Q = 1252;
						break;
					}
					l(I.codepage = Q);
					X = true;
					break;
				case 317:
					I.rrtabid = J;
					break;
				case 25:
					I.winlocked = J;
					break;
				case 439:
					t.opts["RefreshAll"] = J;
					break;
				case 12:
					t.opts["CalcCount"] = J;
					break;
				case 16:
					t.opts["CalcDelta"] = J;
					break;
				case 17:
					t.opts["CalcIter"] = J;
					break;
				case 13:
					t.opts["CalcMode"] = J;
					break;
				case 14:
					t.opts["CalcPrecision"] = J;
					break;
				case 95:
					t.opts["CalcSaveRecalc"] = J;
					break;
				case 15:
					I.CalcRefMode = J;
					break;
				case 2211:
					t.opts.FullCalc = J;
					break;
				case 129:
					if (J.fDialog) n["!type"] = "dialog";
					if (!J.fBelow)(n["!outline"] || (n["!outline"] = {})).above = true;
					if (!J.fRight)(n["!outline"] || (n["!outline"] = {})).left = true;
					break;
				case 67:
					;
				case 579:
					;
				case 1091:
					;
				case 224:
					E.push(J);
					break;
				case 430:
					U.push([J]);
					U[U.length - 1].XTI = [];
					break;
				case 35:
					;
				case 547:
					U[U.length - 1].push(J);
					break;
				case 24:
					;
				case 536:
					$ = {
						Name: J.Name,
						Ref: Md(J.rgce, s, null, U, I)
					};
					if (J.itab > 0) $.Sheet = J.itab - 1;
					U.names.push($);
					if (!U[0]) {
						U[0] = [];
						U[0].XTI = []
					}
					U[U.length - 1].push(J);
					if (J.Name == "_xlnm._FilterDatabase" && J.itab > 0) if (J.rgce && J.rgce[0] && J.rgce[0][0] && J.rgce[0][0][0] == "PtgArea3d") V[J.itab - 1] = {
						ref: Va(J.rgce[0][0][1][2])
					};
					break;
				case 22:
					I.ExternCount = J;
					break;
				case 23:
					if (U.length == 0) {
						U[0] = [];
						U[0].XTI = []
					}
					U[U.length - 1].XTI = U[U.length - 1].XTI.concat(J);
					U.XTI = U.XTI.concat(J);
					break;
				case 2196:
					if (I.biff < 8) break;
					if ($ != null) $.Comment = J[1];
					break;
				case 18:
					n["!protect"] = J;
					break;
				case 19:
					if (J !== 0 && I.WTF) console.error("Password verifier: " + J);
					break;
				case 133:
					{
						i[I.biff == 4 ? I.snames.length: J.pos] = J;
						I.snames.push(J.name)
					}
					break;
				case 10:
					{
						if (--W ? !A: A) break;
						if (s.e) {
							if (s.e.r > 0 && s.e.c > 0) {
								s.e.r--;
								s.e.c--;
								n["!ref"] = Va(s);
								if (r.sheetRows && r.sheetRows <= s.e.r) {
									var ee = s.e.r;
									s.e.r = r.sheetRows - 1;
									n["!fullref"] = n["!ref"];
									n["!ref"] = Va(s);
									s.e.r = ee
								}
								s.e.r++;
								s.e.c++
							}
							if (F.length > 0) n["!merges"] = F;
							if (D.length > 0) n["!objects"] = D;
							if (P.length > 0) n["!cols"] = P;
							if (L.length > 0) n["!rows"] = L;
							S.Sheets.push(x)
						}
						if (c === "") u = n;
						else a[c] = n;
						n = {};
						if (r.dense) n["!data"] = []
					}
					break;
				case 9:
					;
				case 521:
					;
				case 1033:
					;
				case 2057:
					{
						if (I.biff === 8) I.biff = {
							9 : 2,
							521 : 3,
							1033 : 4
						} [j] || {
							512 : 2,
							768 : 3,
							1024 : 4,
							1280 : 5,
							1536 : 8,
							2 : 2,
							7 : 2
						} [J.BIFFVer] || 8;
						I.biffguess = J.BIFFVer == 0;
						if (J.BIFFVer == 0 && J.dt == 4096) {
							I.biff = 5;
							X = true;
							l(I.codepage = 28591)
						}
						if (I.biff == 4 && J.dt & 256) A = true;
						if (I.biff == 8 && J.BIFFVer == 0 && J.dt == 16) I.biff = 2;
						if (W++&&!A) break;
						n = {};
						if (r.dense) n["!data"] = [];
						if (I.biff < 8 && !X) {
							X = true;
							l(I.codepage = r.codepage || 1252)
						}
						if (I.biff == 4 && A) {
							c = (i[I.snames.indexOf(c) + 1] || {
								name: ""
							}).name
						} else if (I.biff < 5 || J.BIFFVer == 0 && J.dt == 4096) {
							if (c === "") c = "Sheet1";
							s = {
								s: {
									r: 0,
									c: 0
								},
								e: {
									r: 0,
									c: 0
								}
							};
							var re = {
								pos: e.l - K,
								name: c
							};
							i[re.pos] = re;
							I.snames.push(c)
						} else c = (i[G] || {
							name: ""
						}).name;
						if (J.dt == 32) n["!type"] = "chart";
						if (J.dt == 64) n["!type"] = "macro";
						F = [];
						D = [];
						I.arrayf = k = [];
						P = [];
						L = [];
						M = false;
						x = {
							Hidden: (i[G] || {
								hs: 0
							}).hs,
							name: c
						}
					}
					break;
				case 515:
					;
				case 3:
					;
				case 2:
					{
						if (n["!type"] == "chart") if (r.dense ? (n["!data"][J.r] || [])[J.c] : n[La(J.c) + Na(J.r)])++J.c;
						T = {
							ixfe: J.ixfe,
							XF: E[J.ixfe] || {},
							v: J.val,
							t: "n"
						};
						if (z > 0) T.z = T.XF && T.XF.numFmtId && H[T.XF.numFmtId] || H[T.ixfe >> 8 & 63];
						Xb(T, r, t.opts.Date1904);
						O({
							c: J.c,
							r: J.r
						},
						T, r)
					}
					break;
				case 5:
					;
				case 517:
					{
						T = {
							ixfe: J.ixfe,
							XF: E[J.ixfe],
							v: J.val,
							t: J.t
						};
						if (z > 0) T.z = T.XF && T.XF.numFmtId && H[T.XF.numFmtId] || H[T.ixfe >> 8 & 63];
						Xb(T, r, t.opts.Date1904);
						O({
							c: J.c,
							r: J.r
						},
						T, r)
					}
					break;
				case 638:
					{
						T = {
							ixfe: J.ixfe,
							XF: E[J.ixfe],
							v: J.rknum,
							t: "n"
						};
						if (z > 0) T.z = T.XF && T.XF.numFmtId && H[T.XF.numFmtId] || H[T.ixfe >> 8 & 63];
						Xb(T, r, t.opts.Date1904);
						O({
							c: J.c,
							r: J.r
						},
						T, r)
					}
					break;
				case 189:
					{
						for (var te = J.c; te <= J.C; ++te) {
							var ae = J.rkrec[te - J.c][0];
							T = {
								ixfe: ae,
								XF: E[ae],
								v: J.rkrec[te - J.c][1],
								t: "n"
							};
							if (z > 0) T.z = T.XF && T.XF.numFmtId && H[T.XF.numFmtId] || H[T.ixfe >> 8 & 63];
							Xb(T, r, t.opts.Date1904);
							O({
								c: te,
								r: J.r
							},
							T, r)
						}
					}
					break;
				case 6:
					;
				case 518:
					;
				case 1030:
					{
						if (J.val == "String") {
							f = J;
							break
						}
						T = Gb(J.val, J.cell.ixfe, J.tt);
						T.XF = E[T.ixfe];
						if (r.cellFormula) {
							var ne = J.formula;
							if (ne && ne[0] && ne[0][0] && ne[0][0][0] == "PtgExp") {
								var ie = ne[0][0][1][0],
								se = ne[0][0][1][1];
								var fe = za({
									r: ie,
									c: se
								});
								if (w[fe]) T.f = "" + Md(J.formula, s, J.cell, U, I);
								else T.F = ((r.dense ? (n["!data"][ie] || [])[se] : n[fe]) || {}).F
							} else T.f = "" + Md(J.formula, s, J.cell, U, I)
						}
						if (z > 0) T.z = T.XF && T.XF.numFmtId && H[T.XF.numFmtId] || H[T.ixfe >> 8 & 63];
						Xb(T, r, t.opts.Date1904);
						O(J.cell, T, r);
						f = J
					}
					break;
				case 7:
					;
				case 519:
					{
						if (f) {
							f.val = J;
							T = Gb(J, f.cell.ixfe, "s");
							T.XF = E[T.ixfe];
							if (r.cellFormula) {
								T.f = "" + Md(f.formula, s, f.cell, U, I)
							}
							if (z > 0) T.z = T.XF && T.XF.numFmtId && H[T.XF.numFmtId] || H[T.ixfe >> 8 & 63];
							Xb(T, r, t.opts.Date1904);
							O(f.cell, T, r);
							f = null
						} else throw new Error("String record expects Formula")
					}
					break;
				case 33:
					;
				case 545:
					{
						k.push(J);
						var le = za(J[0].s);
						v = r.dense ? (n["!data"][J[0].s.r] || [])[J[0].s.c] : n[le];
						if (r.cellFormula && v) {
							if (!f) break;
							if (!le || !v) break;
							v.f = "" + Md(J[1], s, J[0], U, I);
							v.F = Va(J[0])
						}
					}
					break;
				case 1212:
					{
						if (!r.cellFormula) break;
						if (d) {
							if (!f) break;
							w[za(f.cell)] = J[0];
							v = r.dense ? (n["!data"][f.cell.r] || [])[f.cell.c] : n[za(f.cell)]; (v || {}).f = "" + Md(J[0], s, h, U, I)
						}
					}
					break;
				case 253:
					T = Gb(o[J.isst].t, J.ixfe, "s");
					if (o[J.isst].h) T.h = o[J.isst].h;
					T.XF = E[T.ixfe];
					if (z > 0) T.z = T.XF && T.XF.numFmtId && H[T.XF.numFmtId] || H[T.ixfe >> 8 & 63];
					Xb(T, r, t.opts.Date1904);
					O({
						c: J.c,
						r: J.r
					},
					T, r);
					break;
				case 513:
					if (r.sheetStubs) {
						T = {
							ixfe: J.ixfe,
							XF: E[J.ixfe],
							t: "z"
						};
						if (z > 0) T.z = T.XF && T.XF.numFmtId && H[T.XF.numFmtId] || H[T.ixfe >> 8 & 63];
						Xb(T, r, t.opts.Date1904);
						O({
							c: J.c,
							r: J.r
						},
						T, r)
					}
					break;
				case 190:
					if (r.sheetStubs) {
						for (var oe = J.c; oe <= J.C; ++oe) {
							var ce = J.ixfe[oe - J.c];
							T = {
								ixfe: ce,
								XF: E[ce],
								t: "z"
							};
							if (z > 0) T.z = T.XF && T.XF.numFmtId && H[T.XF.numFmtId] || H[T.ixfe >> 8 & 63];
							Xb(T, r, t.opts.Date1904);
							O({
								c: oe,
								r: J.r
							},
							T, r)
						}
					}
					break;
				case 214:
					;
				case 516:
					;
				case 4:
					T = Gb(J.val, J.ixfe, "s");
					T.XF = E[T.ixfe];
					if (z > 0) T.z = T.XF && T.XF.numFmtId && H[T.XF.numFmtId] || H[T.ixfe >> 8 & 63];
					Xb(T, r, t.opts.Date1904);
					O({
						c: J.c,
						r: J.r
					},
					T, r);
					break;
				case 0:
					;
				case 512:
					{
						if (W === 1) s = J
					}
					break;
				case 252:
					{
						o = J
					}
					break;
				case 1054:
					{
						if (I.biff >= 3 && I.biff <= 4) {
							H[z++] = J[1];
							for (var ue = 0; ue < z + 163; ++ue) if (q[ue] == J[1]) break;
							if (ue >= 163) Je(J[1], z + 163)
						} else Je(J[1], J[0])
					}
					break;
				case 30:
					{
						H[z++] = J;
						for (var he = 0; he < z + 163; ++he) if (q[he] == J) break;
						if (he >= 163) Je(J, z + 163)
					}
					break;
				case 229:
					F = F.concat(J);
					break;
				case 93:
					D[J.cmo[0]] = I.lastobj = J;
					break;
				case 438:
					I.lastobj.TxO = J;
					break;
				case 127:
					I.lastobj.ImData = J;
					break;
				case 440:
					{
						for (b = J[0].s.r; b <= J[0].e.r; ++b) for (m = J[0].s.c; m <= J[0].e.c; ++m) {
							v = r.dense ? (n["!data"][b] || [])[m] : n[za({
								c: m,
								r: b
							})];
							if (v) v.l = J[1]
						}
					}
					break;
				case 2048:
					{
						for (b = J[0].s.r; b <= J[0].e.r; ++b) for (m = J[0].s.c; m <= J[0].e.c; ++m) {
							v = r.dense ? (n["!data"][b] || [])[m] : n[za({
								c: m,
								r: b
							})];
							if (v && v.l) v.l.Tooltip = J[1]
						}
					}
					break;
				case 28:
					{
						v = r.dense ? (n["!data"][J[0].r] || [])[J[0].c] : n[za(J[0])];
						if (!v) {
							if (r.dense) {
								if (!n["!data"][J[0].r]) n["!data"][J[0].r] = [];
								v = n["!data"][J[0].r][J[0].c] = {
									t: "z"
								}
							} else {
								v = n[za(J[0])] = {
									t: "z"
								}
							}
							s.e.r = Math.max(s.e.r, J[0].r);
							s.s.r = Math.min(s.s.r, J[0].r);
							s.e.c = Math.max(s.e.c, J[0].c);
							s.s.c = Math.min(s.s.c, J[0].c)
						}
						if (!v.c) v.c = [];
						if (I.biff <= 5 && I.biff >= 2) p = {
							a: "SheetJ5",
							t: J[1]
						};
						else {
							var de = D[J[2]];
							p = {
								a: J[1],
								t: de.TxO.t
							};
							if (J[3] != null && !(J[3] & 2)) v.c.hidden = true
						}
						v.c.push(p)
					}
					break;
				case 2173:
					hu(E[J.ixfe], J.ext);
					break;
				case 125:
					{
						if (!I.cellStyles) break;
						while (J.e >= J.s) {
							P[J.e--] = {
								width: J.w / 256,
								level: J.level || 0,
								hidden: !!(J.flags & 1)
							};
							if (!M) {
								M = true;
								ac(J.w / 256)
							}
							nc(P[J.e + 1])
						}
					}
					break;
				case 520:
					{
						var ve = {};
						if (J.level != null) {
							L[J.r] = ve;
							ve.level = J.level
						}
						if (J.hidden) {
							L[J.r] = ve;
							ve.hidden = true
						}
						if (J.hpt) {
							L[J.r] = ve;
							ve.hpt = J.hpt;
							ve.hpx = lc(J.hpt)
						}
					}
					break;
				case 38:
					;
				case 39:
					;
				case 40:
					;
				case 41:
					if (!n["!margins"]) Tv(n["!margins"] = {});
					n["!margins"][{
						38 : "left",
						39 : "right",
						40 : "top",
						41 : "bottom"
					} [j]] = J;
					break;
				case 161:
					if (!n["!margins"]) Tv(n["!margins"] = {});
					n["!margins"].header = J.header;
					n["!margins"].footer = J.footer;
					break;
				case 574:
					if (J.RTL) S.Views[0].RTL = true;
					break;
				case 146:
					_ = J;
					break;
				case 2198:
					N = J;
					break;
				case 140:
					y = J;
					break;
				case 442:
					{
						if (!c) S.WBProps.CodeName = J || "ThisWorkbook";
						else x.CodeName = J || x.name
					}
					break;
				}
			} else {
				if (!Y) console.error("Missing Info for XLS Record 0x" + j.toString(16));
				e.l += K
			}
		}
		t.SheetNames = ir(i).sort(function(e, r) {
			return Number(e) - Number(r)
		}).map(function(e) {
			return i[e].name
		});
		if (!r.bookSheets) t.Sheets = a;
		if (!t.SheetNames.length && u["!ref"]) {
			t.SheetNames.push("Sheet1");
			if (t.Sheets) t.Sheets["Sheet1"] = u
		} else t.Preamble = u;
		if (t.Sheets) V.forEach(function(e, r) {
			t.Sheets[t.SheetNames[r]]["!autofilter"] = e
		});
		t.Strings = o;
		t.SSF = Tr(q);
		if (I.enc) t.Encryption = I.enc;
		if (N) t.Themes = N;
		t.Metadata = {};
		if (y !== undefined) t.Metadata.Country = y;
		if (U.names.length > 0) S.Names = U.names;
		t.Workbook = S;
		return t
	}
	var Kb = {
		SI: "e0859ff2f94f6810ab9108002b27b3d9",
		DSI: "02d5cdd59c2e1b10939708002b2cf9ae",
		UDI: "05d5cdd59c2e1b10939708002b2cf9ae"
	};
	function Yb(e, r, t) {
		var a = Qe.find(e, "/!DocumentSummaryInformation");
		if (a && a.size > 0) try {
			var n = ss(a, Yn, Kb.DSI);
			for (var i in n) r[i] = n[i]
		} catch(s) {
			if (t.WTF) throw s
		}
		var f = Qe.find(e, "/!SummaryInformation");
		if (f && f.size > 0) try {
			var l = ss(f, Zn, Kb.SI);
			for (var o in l) if (r[o] == null) r[o] = l[o]
		} catch(s) {
			if (t.WTF) throw s
		}
		if (r.HeadingPairs && r.TitlesOfParts) {
			Ri(r.HeadingPairs, r.TitlesOfParts, r, t);
			delete r.HeadingPairs;
			delete r.TitlesOfParts
		}
	}
	function Zb(e, r) {
		var t = [],
		a = [],
		n = [];
		var i = 0,
		s;
		var f = sr(Yn, "n");
		var l = sr(Zn, "n");
		if (e.Props) {
			s = ir(e.Props);
			for (i = 0; i < s.length; ++i)(Object.prototype.hasOwnProperty.call(f, s[i]) ? t: Object.prototype.hasOwnProperty.call(l, s[i]) ? a: n).push([s[i], e.Props[s[i]]])
		}
		if (e.Custprops) {
			s = ir(e.Custprops);
			for (i = 0; i < s.length; ++i) if (!Object.prototype.hasOwnProperty.call(e.Props || {},
			s[i]))(Object.prototype.hasOwnProperty.call(f, s[i]) ? t: Object.prototype.hasOwnProperty.call(l, s[i]) ? a: n).push([s[i], e.Custprops[s[i]]])
		}
		var o = [];
		for (i = 0; i < n.length; ++i) {
			if (as.indexOf(n[i][0]) > -1 || Ci.indexOf(n[i][0]) > -1) continue;
			if (n[i][1] == null) continue;
			o.push(n[i])
		}
		if (a.length) Qe.utils.cfb_add(r, "/SummaryInformation", fs(a, Kb.SI, l, Zn));
		if (t.length || o.length) Qe.utils.cfb_add(r, "/DocumentSummaryInformation", fs(t, Kb.DSI, f, Yn, o.length ? o: null, Kb.UDI))
	}
	function Jb(e, r) {
		if (!r) r = {};
		ek(r);
		o();
		if (r.codepage) s(r.codepage);
		var t, a;
		if (e.FullPaths) {
			if (Qe.find(e, "/encryption")) throw new Error("File is password-protected");
			t = Qe.find(e, "!CompObj");
			a = Qe.find(e, "/Workbook") || Qe.find(e, "/Book")
		} else {
			switch (r.type) {
			case "base64":
				e = O(_(e));
				break;
			case "binary":
				e = O(e);
				break;
			case "buffer":
				break;
			case "array":
				if (!Array.isArray(e)) e = Array.prototype.slice.call(e);
				break;
			}
			Ta(e, 0);
			a = {
				content: e
			}
		}
		var n;
		var i;
		if (t) Hb(t);
		if (r.bookProps && !r.bookSheets) n = {};
		else {
			var f = S ? "buffer": "array";
			if (a && a.content) n = jb(a.content, r);
			else if ((i = Qe.find(e, "PerfectOffice_MAIN")) && i.content) n = Jl.to_workbook(i.content, (r.type = f, r));
			else if ((i = Qe.find(e, "NativeContent_MAIN")) && i.content) n = Jl.to_workbook(i.content, (r.type = f, r));
			else if ((i = Qe.find(e, "MN0")) && i.content) throw new Error("Unsupported Works 4 for Mac file");
			else throw new Error("Cannot find Workbook stream");
			if (r.bookVBA && e.FullPaths && Qe.find(e, "/_VBA_PROJECT_CUR/VBA/dir")) n.vbaraw = Ku(e)
		}
		var l = {};
		if (e.FullPaths) Yb(e, l, r);
		n.Props = n.Custprops = l;
		if (r.bookFiles) n.cfb = e;
		return n
	}
	function qb(e, r) {
		var t = r || {};
		var a = Qe.utils.cfb_new({
			root: "R"
		});
		var n = "/Workbook";
		switch (t.bookType || "xls") {
		case "xls":
			t.bookType = "biff8";
		case "xla":
			if (!t.bookType) t.bookType = "xla";
		case "biff8":
			n = "/Workbook";
			t.biff = 8;
			break;
		case "biff5":
			n = "/Book";
			t.biff = 5;
			break;
		default:
			throw new Error("invalid type " + t.bookType + " for XLS CFB");
		}
		Qe.utils.cfb_add(a, n, Eg(e, t));
		if (t.biff == 8 && (e.Props || e.Custprops)) Zb(e, a);
		if (t.biff == 8 && e.vbaraw) Yu(a, Qe.read(e.vbaraw, {
			type: typeof e.vbaraw == "string" ? "binary": "buffer"
		}));
		return a
	}
	var Qb = {
		0 : {
			f: np
		},
		1 : {
			f: hp
		},
		2 : {
			f: Np
		},
		3 : {
			f: kp
		},
		4 : {
			f: mp
		},
		5 : {
			f: Cp
		},
		6 : {
			f: Mp
		},
		7 : {
			f: _p
		},
		8 : {
			f: $p
		},
		9 : {
			f: Vp
		},
		10 : {
			f: zp
		},
		11 : {
			f: Hp
		},
		12 : {
			f: vp
		},
		13 : {
			f: Dp
		},
		14 : {
			f: yp
		},
		15 : {
			f: gp
		},
		16 : {
			f: Op
		},
		17 : {
			f: Bp
		},
		18 : {
			f: xp
		},
		19 : {
			f: sn
		},
		20 : {},
		21 : {},
		22 : {},
		23 : {},
		24 : {},
		25 : {},
		26 : {},
		27 : {},
		28 : {},
		29 : {},
		30 : {},
		31 : {},
		32 : {},
		33 : {},
		34 : {},
		35 : {
			T: 1
		},
		36 : {
			T: -1
		},
		37 : {
			T: 1
		},
		38 : {
			T: -1
		},
		39 : {
			f: jm
		},
		40 : {},
		42 : {},
		43 : {
			f: Ec
		},
		44 : {
			f: Tc
		},
		45 : {
			f: Ac
		},
		46 : {
			f: Nc
		},
		47 : {
			f: Rc
		},
		48 : {},
		49 : {
			f: Qa
		},
		50 : {},
		51 : {
			f: pu
		},
		52 : {
			T: 1
		},
		53 : {
			T: -1
		},
		54 : {
			T: 1
		},
		55 : {
			T: -1
		},
		56 : {
			T: 1
		},
		57 : {
			T: -1
		},
		58 : {},
		59 : {},
		60 : {
			f: _l
		},
		62 : {
			f: Lp
		},
		63 : {
			f: Su
		},
		64 : {
			f: fm
		},
		65 : {},
		66 : {},
		67 : {},
		68 : {},
		69 : {},
		70 : {},
		128 : {},
		129 : {
			T: 1
		},
		130 : {
			T: -1
		},
		131 : {
			T: 1,
			f: ya,
			p: 0
		},
		132 : {
			T: -1
		},
		133 : {
			T: 1
		},
		134 : {
			T: -1
		},
		135 : {
			T: 1
		},
		136 : {
			T: -1
		},
		137 : {
			T: 1,
			f: am
		},
		138 : {
			T: -1
		},
		139 : {
			T: 1
		},
		140 : {
			T: -1
		},
		141 : {
			T: 1
		},
		142 : {
			T: -1
		},
		143 : {
			T: 1
		},
		144 : {
			T: -1
		},
		145 : {
			T: 1
		},
		146 : {
			T: -1
		},
		147 : {
			f: cp
		},
		148 : {
			f: fp,
			p: 16
		},
		151 : {
			f: Zp
		},
		152 : {},
		153 : {
			f: $m
		},
		154 : {},
		155 : {},
		156 : {
			f: Hm
		},
		157 : {},
		158 : {},
		159 : {
			T: 1,
			f: uo
		},
		160 : {
			T: -1
		},
		161 : {
			T: 1,
			f: Sn
		},
		162 : {
			T: -1
		},
		163 : {
			T: 1
		},
		164 : {
			T: -1
		},
		165 : {
			T: 1
		},
		166 : {
			T: -1
		},
		167 : {},
		168 : {},
		169 : {},
		170 : {},
		171 : {},
		172 : {
			T: 1
		},
		173 : {
			T: -1
		},
		174 : {},
		175 : {},
		176 : {
			f: Xp
		},
		177 : {
			T: 1
		},
		178 : {
			T: -1
		},
		179 : {
			T: 1
		},
		180 : {
			T: -1
		},
		181 : {
			T: 1
		},
		182 : {
			T: -1
		},
		183 : {
			T: 1
		},
		184 : {
			T: -1
		},
		185 : {
			T: 1
		},
		186 : {
			T: -1
		},
		187 : {
			T: 1
		},
		188 : {
			T: -1
		},
		189 : {
			T: 1
		},
		190 : {
			T: -1
		},
		191 : {
			T: 1
		},
		192 : {
			T: -1
		},
		193 : {
			T: 1
		},
		194 : {
			T: -1
		},
		195 : {
			T: 1
		},
		196 : {
			T: -1
		},
		197 : {
			T: 1
		},
		198 : {
			T: -1
		},
		199 : {
			T: 1
		},
		200 : {
			T: -1
		},
		201 : {
			T: 1
		},
		202 : {
			T: -1
		},
		203 : {
			T: 1
		},
		204 : {
			T: -1
		},
		205 : {
			T: 1
		},
		206 : {
			T: -1
		},
		207 : {
			T: 1
		},
		208 : {
			T: -1
		},
		209 : {
			T: 1
		},
		210 : {
			T: -1
		},
		211 : {
			T: 1
		},
		212 : {
			T: -1
		},
		213 : {
			T: 1
		},
		214 : {
			T: -1
		},
		215 : {
			T: 1
		},
		216 : {
			T: -1
		},
		217 : {
			T: 1
		},
		218 : {
			T: -1
		},
		219 : {
			T: 1
		},
		220 : {
			T: -1
		},
		221 : {
			T: 1
		},
		222 : {
			T: -1
		},
		223 : {
			T: 1
		},
		224 : {
			T: -1
		},
		225 : {
			T: 1
		},
		226 : {
			T: -1
		},
		227 : {
			T: 1
		},
		228 : {
			T: -1
		},
		229 : {
			T: 1
		},
		230 : {
			T: -1
		},
		231 : {
			T: 1
		},
		232 : {
			T: -1
		},
		233 : {
			T: 1
		},
		234 : {
			T: -1
		},
		235 : {
			T: 1
		},
		236 : {
			T: -1
		},
		237 : {
			T: 1
		},
		238 : {
			T: -1
		},
		239 : {
			T: 1
		},
		240 : {
			T: -1
		},
		241 : {
			T: 1
		},
		242 : {
			T: -1
		},
		243 : {
			T: 1
		},
		244 : {
			T: -1
		},
		245 : {
			T: 1
		},
		246 : {
			T: -1
		},
		247 : {
			T: 1
		},
		248 : {
			T: -1
		},
		249 : {
			T: 1
		},
		250 : {
			T: -1
		},
		251 : {
			T: 1
		},
		252 : {
			T: -1
		},
		253 : {
			T: 1
		},
		254 : {
			T: -1
		},
		255 : {
			T: 1
		},
		256 : {
			T: -1
		},
		257 : {
			T: 1
		},
		258 : {
			T: -1
		},
		259 : {
			T: 1
		},
		260 : {
			T: -1
		},
		261 : {
			T: 1
		},
		262 : {
			T: -1
		},
		263 : {
			T: 1
		},
		264 : {
			T: -1
		},
		265 : {
			T: 1
		},
		266 : {
			T: -1
		},
		267 : {
			T: 1
		},
		268 : {
			T: -1
		},
		269 : {
			T: 1
		},
		270 : {
			T: -1
		},
		271 : {
			T: 1
		},
		272 : {
			T: -1
		},
		273 : {
			T: 1
		},
		274 : {
			T: -1
		},
		275 : {
			T: 1
		},
		276 : {
			T: -1
		},
		277 : {},
		278 : {
			T: 1
		},
		279 : {
			T: -1
		},
		280 : {
			T: 1
		},
		281 : {
			T: -1
		},
		282 : {
			T: 1
		},
		283 : {
			T: 1
		},
		284 : {
			T: -1
		},
		285 : {
			T: 1
		},
		286 : {
			T: -1
		},
		287 : {
			T: 1
		},
		288 : {
			T: -1
		},
		289 : {
			T: 1
		},
		290 : {
			T: -1
		},
		291 : {
			T: 1
		},
		292 : {
			T: -1
		},
		293 : {
			T: 1
		},
		294 : {
			T: -1
		},
		295 : {
			T: 1
		},
		296 : {
			T: -1
		},
		297 : {
			T: 1
		},
		298 : {
			T: -1
		},
		299 : {
			T: 1
		},
		300 : {
			T: -1
		},
		301 : {
			T: 1
		},
		302 : {
			T: -1
		},
		303 : {
			T: 1
		},
		304 : {
			T: -1
		},
		305 : {
			T: 1
		},
		306 : {
			T: -1
		},
		307 : {
			T: 1
		},
		308 : {
			T: -1
		},
		309 : {
			T: 1
		},
		310 : {
			T: -1
		},
		311 : {
			T: 1
		},
		312 : {
			T: -1
		},
		313 : {
			T: -1
		},
		314 : {
			T: 1
		},
		315 : {
			T: -1
		},
		316 : {
			T: 1
		},
		317 : {
			T: -1
		},
		318 : {
			T: 1
		},
		319 : {
			T: -1
		},
		320 : {
			T: 1
		},
		321 : {
			T: -1
		},
		322 : {
			T: 1
		},
		323 : {
			T: -1
		},
		324 : {
			T: 1
		},
		325 : {
			T: -1
		},
		326 : {
			T: 1
		},
		327 : {
			T: -1
		},
		328 : {
			T: 1
		},
		329 : {
			T: -1
		},
		330 : {
			T: 1
		},
		331 : {
			T: -1
		},
		332 : {
			T: 1
		},
		333 : {
			T: -1
		},
		334 : {
			T: 1
		},
		335 : {
			f: du
		},
		336 : {
			T: -1
		},
		337 : {
			f: gu,
			T: 1
		},
		338 : {
			T: -1
		},
		339 : {
			T: 1
		},
		340 : {
			T: -1
		},
		341 : {
			T: 1
		},
		342 : {
			T: -1
		},
		343 : {
			T: 1
		},
		344 : {
			T: -1
		},
		345 : {
			T: 1
		},
		346 : {
			T: -1
		},
		347 : {
			T: 1
		},
		348 : {
			T: -1
		},
		349 : {
			T: 1
		},
		350 : {
			T: -1
		},
		351 : {},
		352 : {},
		353 : {
			T: 1
		},
		354 : {
			T: -1
		},
		355 : {
			f: wn
		},
		357 : {},
		358 : {},
		359 : {},
		360 : {
			T: 1
		},
		361 : {},
		362 : {
			f: el
		},
		363 : {},
		364 : {},
		366 : {},
		367 : {},
		368 : {},
		369 : {},
		370 : {},
		371 : {},
		372 : {
			T: 1
		},
		373 : {
			T: -1
		},
		374 : {
			T: 1
		},
		375 : {
			T: -1
		},
		376 : {
			T: 1
		},
		377 : {
			T: -1
		},
		378 : {
			T: 1
		},
		379 : {
			T: -1
		},
		380 : {
			T: 1
		},
		381 : {
			T: -1
		},
		382 : {
			T: 1
		},
		383 : {
			T: -1
		},
		384 : {
			T: 1
		},
		385 : {
			T: -1
		},
		386 : {
			T: 1
		},
		387 : {
			T: -1
		},
		388 : {
			T: 1
		},
		389 : {
			T: -1
		},
		390 : {
			T: 1
		},
		391 : {
			T: -1
		},
		392 : {
			T: 1
		},
		393 : {
			T: -1
		},
		394 : {
			T: 1
		},
		395 : {
			T: -1
		},
		396 : {},
		397 : {},
		398 : {},
		399 : {},
		400 : {},
		401 : {
			T: 1
		},
		403 : {},
		404 : {},
		405 : {},
		406 : {},
		407 : {},
		408 : {},
		409 : {},
		410 : {},
		411 : {},
		412 : {},
		413 : {},
		414 : {},
		415 : {},
		416 : {},
		417 : {},
		418 : {},
		419 : {},
		420 : {},
		421 : {},
		422 : {
			T: 1
		},
		423 : {
			T: 1
		},
		424 : {
			T: -1
		},
		425 : {
			T: -1
		},
		426 : {
			f: Jp
		},
		427 : {
			f: qp
		},
		428 : {},
		429 : {
			T: 1
		},
		430 : {
			T: -1
		},
		431 : {
			T: 1
		},
		432 : {
			T: -1
		},
		433 : {
			T: 1
		},
		434 : {
			T: -1
		},
		435 : {
			T: 1
		},
		436 : {
			T: -1
		},
		437 : {
			T: 1
		},
		438 : {
			T: -1
		},
		439 : {
			T: 1
		},
		440 : {
			T: -1
		},
		441 : {
			T: 1
		},
		442 : {
			T: -1
		},
		443 : {
			T: 1
		},
		444 : {
			T: -1
		},
		445 : {
			T: 1
		},
		446 : {
			T: -1
		},
		447 : {
			T: 1
		},
		448 : {
			T: -1
		},
		449 : {
			T: 1
		},
		450 : {
			T: -1
		},
		451 : {
			T: 1
		},
		452 : {
			T: -1
		},
		453 : {
			T: 1
		},
		454 : {
			T: -1
		},
		455 : {
			T: 1
		},
		456 : {
			T: -1
		},
		457 : {
			T: 1
		},
		458 : {
			T: -1
		},
		459 : {
			T: 1
		},
		460 : {
			T: -1
		},
		461 : {
			T: 1
		},
		462 : {
			T: -1
		},
		463 : {
			T: 1
		},
		464 : {
			T: -1
		},
		465 : {
			T: 1
		},
		466 : {
			T: -1
		},
		467 : {
			T: 1
		},
		468 : {
			T: -1
		},
		469 : {
			T: 1
		},
		470 : {
			T: -1
		},
		471 : {},
		472 : {},
		473 : {
			T: 1
		},
		474 : {
			T: -1
		},
		475 : {},
		476 : {
			f: rm
		},
		477 : {},
		478 : {},
		479 : {
			T: 1
		},
		480 : {
			T: -1
		},
		481 : {
			T: 1
		},
		482 : {
			T: -1
		},
		483 : {
			T: 1
		},
		484 : {
			T: -1
		},
		485 : {
			f: op
		},
		486 : {
			T: 1
		},
		487 : {
			T: -1
		},
		488 : {
			T: 1
		},
		489 : {
			T: -1
		},
		490 : {
			T: 1
		},
		491 : {
			T: -1
		},
		492 : {
			T: 1
		},
		493 : {
			T: -1
		},
		494 : {
			f: Kp
		},
		495 : {
			T: 1
		},
		496 : {
			T: -1
		},
		497 : {
			T: 1
		},
		498 : {
			T: -1
		},
		499 : {},
		500 : {
			T: 1
		},
		501 : {
			T: -1
		},
		502 : {
			T: 1
		},
		503 : {
			T: -1
		},
		504 : {},
		505 : {
			T: 1
		},
		506 : {
			T: -1
		},
		507 : {},
		508 : {
			T: 1
		},
		509 : {
			T: -1
		},
		510 : {
			T: 1
		},
		511 : {
			T: -1
		},
		512 : {},
		513 : {},
		514 : {
			T: 1
		},
		515 : {
			T: -1
		},
		516 : {
			T: 1
		},
		517 : {
			T: -1
		},
		518 : {
			T: 1
		},
		519 : {
			T: -1
		},
		520 : {
			T: 1
		},
		521 : {
			T: -1
		},
		522 : {},
		523 : {},
		524 : {},
		525 : {},
		526 : {},
		527 : {},
		528 : {
			T: 1
		},
		529 : {
			T: -1
		},
		530 : {
			T: 1
		},
		531 : {
			T: -1
		},
		532 : {
			T: 1
		},
		533 : {
			T: -1
		},
		534 : {},
		535 : {},
		536 : {},
		537 : {},
		538 : {
			T: 1
		},
		539 : {
			T: -1
		},
		540 : {
			T: 1
		},
		541 : {
			T: -1
		},
		542 : {
			T: 1
		},
		548 : {},
		549 : {},
		550 : {
			f: wn
		},
		551 : {
			f: mn
		},
		552 : {},
		553 : {},
		554 : {
			T: 1
		},
		555 : {
			T: -1
		},
		556 : {
			T: 1
		},
		557 : {
			T: -1
		},
		558 : {
			T: 1
		},
		559 : {
			T: -1
		},
		560 : {
			T: 1
		},
		561 : {
			T: -1
		},
		562 : {},
		564 : {},
		565 : {
			T: 1
		},
		566 : {
			T: -1
		},
		569 : {
			T: 1
		},
		570 : {
			T: -1
		},
		572 : {},
		573 : {
			T: 1
		},
		574 : {
			T: -1
		},
		577 : {},
		578 : {},
		579 : {},
		580 : {},
		581 : {},
		582 : {},
		583 : {},
		584 : {},
		585 : {},
		586 : {},
		587 : {},
		588 : {
			T: -1
		},
		589 : {},
		590 : {
			T: 1
		},
		591 : {
			T: -1
		},
		592 : {
			T: 1
		},
		593 : {
			T: -1
		},
		594 : {
			T: 1
		},
		595 : {
			T: -1
		},
		596 : {},
		597 : {
			T: 1
		},
		598 : {
			T: -1
		},
		599 : {
			T: 1
		},
		600 : {
			T: -1
		},
		601 : {
			T: 1
		},
		602 : {
			T: -1
		},
		603 : {
			T: 1
		},
		604 : {
			T: -1
		},
		605 : {
			T: 1
		},
		606 : {
			T: -1
		},
		607 : {},
		608 : {
			T: 1
		},
		609 : {
			T: -1
		},
		610 : {},
		611 : {
			T: 1
		},
		612 : {
			T: -1
		},
		613 : {
			T: 1
		},
		614 : {
			T: -1
		},
		615 : {
			T: 1
		},
		616 : {
			T: -1
		},
		617 : {
			T: 1
		},
		618 : {
			T: -1
		},
		619 : {
			T: 1
		},
		620 : {
			T: -1
		},
		625 : {},
		626 : {
			T: 1
		},
		627 : {
			T: -1
		},
		628 : {
			T: 1
		},
		629 : {
			T: -1
		},
		630 : {
			T: 1
		},
		631 : {
			T: -1
		},
		632 : {
			f: Vu
		},
		633 : {
			T: 1
		},
		634 : {
			T: -1
		},
		635 : {
			T: 1,
			f: zu
		},
		636 : {
			T: -1
		},
		637 : {
			f: ln
		},
		638 : {
			T: 1
		},
		639 : {},
		640 : {
			T: -1
		},
		641 : {
			T: 1
		},
		642 : {
			T: -1
		},
		643 : {
			T: 1
		},
		644 : {},
		645 : {
			T: -1
		},
		646 : {
			T: 1
		},
		648 : {
			T: 1
		},
		649 : {},
		650 : {
			T: -1
		},
		651 : {
			f: Sm
		},
		652 : {},
		653 : {
			T: 1
		},
		654 : {
			T: -1
		},
		655 : {
			T: 1
		},
		656 : {
			T: -1
		},
		657 : {
			T: 1
		},
		658 : {
			T: -1
		},
		659 : {},
		660 : {
			T: 1
		},
		661 : {},
		662 : {
			T: -1
		},
		663 : {},
		664 : {
			T: 1
		},
		665 : {},
		666 : {
			T: -1
		},
		667 : {},
		668 : {},
		669 : {},
		671 : {
			T: 1
		},
		672 : {
			T: -1
		},
		673 : {
			T: 1
		},
		674 : {
			T: -1
		},
		675 : {},
		676 : {},
		677 : {},
		678 : {},
		679 : {},
		680 : {},
		681 : {},
		1024 : {},
		1025 : {},
		1026 : {
			T: 1
		},
		1027 : {
			T: -1
		},
		1028 : {
			T: 1
		},
		1029 : {
			T: -1
		},
		1030 : {},
		1031 : {
			T: 1
		},
		1032 : {
			T: -1
		},
		1033 : {
			T: 1
		},
		1034 : {
			T: -1
		},
		1035 : {},
		1036 : {},
		1037 : {},
		1038 : {
			T: 1
		},
		1039 : {
			T: -1
		},
		1040 : {},
		1041 : {
			T: 1
		},
		1042 : {
			T: -1
		},
		1043 : {},
		1044 : {},
		1045 : {},
		1046 : {
			T: 1
		},
		1047 : {
			T: -1
		},
		1048 : {
			T: 1
		},
		1049 : {
			T: -1
		},
		1050 : {},
		1051 : {
			T: 1
		},
		1052 : {
			T: 1
		},
		1053 : {
			f: lm
		},
		1054 : {
			T: 1
		},
		1055 : {},
		1056 : {
			T: 1
		},
		1057 : {
			T: -1
		},
		1058 : {
			T: 1
		},
		1059 : {
			T: -1
		},
		1061 : {},
		1062 : {
			T: 1
		},
		1063 : {
			T: -1
		},
		1064 : {
			T: 1
		},
		1065 : {
			T: -1
		},
		1066 : {
			T: 1
		},
		1067 : {
			T: -1
		},
		1068 : {
			T: 1
		},
		1069 : {
			T: -1
		},
		1070 : {
			T: 1
		},
		1071 : {
			T: -1
		},
		1072 : {
			T: 1
		},
		1073 : {
			T: -1
		},
		1075 : {
			T: 1
		},
		1076 : {
			T: -1
		},
		1077 : {
			T: 1
		},
		1078 : {
			T: -1
		},
		1079 : {
			T: 1
		},
		1080 : {
			T: -1
		},
		1081 : {
			T: 1
		},
		1082 : {
			T: -1
		},
		1083 : {
			T: 1
		},
		1084 : {
			T: -1
		},
		1085 : {},
		1086 : {
			T: 1
		},
		1087 : {
			T: -1
		},
		1088 : {
			T: 1
		},
		1089 : {
			T: -1
		},
		1090 : {
			T: 1
		},
		1091 : {
			T: -1
		},
		1092 : {
			T: 1
		},
		1093 : {
			T: -1
		},
		1094 : {
			T: 1
		},
		1095 : {
			T: -1
		},
		1096 : {},
		1097 : {
			T: 1
		},
		1098 : {},
		1099 : {
			T: -1
		},
		1100 : {
			T: 1
		},
		1101 : {
			T: -1
		},
		1102 : {},
		1103 : {},
		1104 : {},
		1105 : {},
		1111 : {},
		1112 : {},
		1113 : {
			T: 1
		},
		1114 : {
			T: -1
		},
		1115 : {
			T: 1
		},
		1116 : {
			T: -1
		},
		1117 : {},
		1118 : {
			T: 1
		},
		1119 : {
			T: -1
		},
		1120 : {
			T: 1
		},
		1121 : {
			T: -1
		},
		1122 : {
			T: 1
		},
		1123 : {
			T: -1
		},
		1124 : {
			T: 1
		},
		1125 : {
			T: -1
		},
		1126 : {},
		1128 : {
			T: 1
		},
		1129 : {
			T: -1
		},
		1130 : {},
		1131 : {
			T: 1
		},
		1132 : {
			T: -1
		},
		1133 : {
			T: 1
		},
		1134 : {
			T: -1
		},
		1135 : {
			T: 1
		},
		1136 : {
			T: -1
		},
		1137 : {
			T: 1
		},
		1138 : {
			T: -1
		},
		1139 : {
			T: 1
		},
		1140 : {
			T: -1
		},
		1141 : {},
		1142 : {
			T: 1
		},
		1143 : {
			T: -1
		},
		1144 : {
			T: 1
		},
		1145 : {
			T: -1
		},
		1146 : {},
		1147 : {
			T: 1
		},
		1148 : {
			T: -1
		},
		1149 : {
			T: 1
		},
		1150 : {
			T: -1
		},
		1152 : {
			T: 1
		},
		1153 : {
			T: -1
		},
		1154 : {
			T: -1
		},
		1155 : {
			T: -1
		},
		1156 : {
			T: -1
		},
		1157 : {
			T: 1
		},
		1158 : {
			T: -1
		},
		1159 : {
			T: 1
		},
		1160 : {
			T: -1
		},
		1161 : {
			T: 1
		},
		1162 : {
			T: -1
		},
		1163 : {
			T: 1
		},
		1164 : {
			T: -1
		},
		1165 : {
			T: 1
		},
		1166 : {
			T: -1
		},
		1167 : {
			T: 1
		},
		1168 : {
			T: -1
		},
		1169 : {
			T: 1
		},
		1170 : {
			T: -1
		},
		1171 : {},
		1172 : {
			T: 1
		},
		1173 : {
			T: -1
		},
		1177 : {},
		1178 : {
			T: 1
		},
		1180 : {},
		1181 : {},
		1182 : {},
		2048 : {
			T: 1
		},
		2049 : {
			T: -1
		},
		2050 : {},
		2051 : {
			T: 1
		},
		2052 : {
			T: -1
		},
		2053 : {},
		2054 : {},
		2055 : {
			T: 1
		},
		2056 : {
			T: -1
		},
		2057 : {
			T: 1
		},
		2058 : {
			T: -1
		},
		2060 : {},
		2067 : {},
		2068 : {
			T: 1
		},
		2069 : {
			T: -1
		},
		2070 : {},
		2071 : {},
		2072 : {
			T: 1
		},
		2073 : {
			T: -1
		},
		2075 : {},
		2076 : {},
		2077 : {
			T: 1
		},
		2078 : {
			T: -1
		},
		2079 : {},
		2080 : {
			T: 1
		},
		2081 : {
			T: -1
		},
		2082 : {},
		2083 : {
			T: 1
		},
		2084 : {
			T: -1
		},
		2085 : {
			T: 1
		},
		2086 : {
			T: -1
		},
		2087 : {
			T: 1
		},
		2088 : {
			T: -1
		},
		2089 : {
			T: 1
		},
		2090 : {
			T: -1
		},
		2091 : {},
		2092 : {},
		2093 : {
			T: 1
		},
		2094 : {
			T: -1
		},
		2095 : {},
		2096 : {
			T: 1
		},
		2097 : {
			T: -1
		},
		2098 : {
			T: 1
		},
		2099 : {
			T: -1
		},
		2100 : {
			T: 1
		},
		2101 : {
			T: -1
		},
		2102 : {},
		2103 : {
			T: 1
		},
		2104 : {
			T: -1
		},
		2105 : {},
		2106 : {
			T: 1
		},
		2107 : {
			T: -1
		},
		2108 : {},
		2109 : {
			T: 1
		},
		2110 : {
			T: -1
		},
		2111 : {
			T: 1
		},
		2112 : {
			T: -1
		},
		2113 : {
			T: 1
		},
		2114 : {
			T: -1
		},
		2115 : {},
		2116 : {},
		2117 : {},
		2118 : {
			T: 1
		},
		2119 : {
			T: -1
		},
		2120 : {},
		2121 : {
			T: 1
		},
		2122 : {
			T: -1
		},
		2123 : {
			T: 1
		},
		2124 : {
			T: -1
		},
		2125 : {},
		2126 : {
			T: 1
		},
		2127 : {
			T: -1
		},
		2128 : {},
		2129 : {
			T: 1
		},
		2130 : {
			T: -1
		},
		2131 : {
			T: 1
		},
		2132 : {
			T: -1
		},
		2133 : {
			T: 1
		},
		2134 : {},
		2135 : {},
		2136 : {},
		2137 : {
			T: 1
		},
		2138 : {
			T: -1
		},
		2139 : {
			T: 1
		},
		2140 : {
			T: -1
		},
		2141 : {},
		3072 : {},
		3073 : {},
		4096 : {
			T: 1
		},
		4097 : {
			T: -1
		},
		5002 : {
			T: 1
		},
		5003 : {
			T: -1
		},
		5081 : {
			T: 1
		},
		5082 : {
			T: -1
		},
		5083 : {},
		5084 : {
			T: 1
		},
		5085 : {
			T: -1
		},
		5086 : {
			T: 1
		},
		5087 : {
			T: -1
		},
		5088 : {},
		5089 : {},
		5090 : {},
		5092 : {
			T: 1
		},
		5093 : {
			T: -1
		},
		5094 : {},
		5095 : {
			T: 1
		},
		5096 : {
			T: -1
		},
		5097 : {},
		5099 : {},
		65535 : {
			n: ""
		}
	};
	var eg = {
		6 : {
			f: $d
		},
		10 : {
			f: ls
		},
		12 : {
			f: ds
		},
		13 : {
			f: ds
		},
		14 : {
			f: us
		},
		15 : {
			f: us
		},
		16 : {
			f: An
		},
		17 : {
			f: us
		},
		18 : {
			f: us
		},
		19 : {
			f: ds
		},
		20 : {
			f: Zf
		},
		21 : {
			f: Zf
		},
		23 : {
			f: el
		},
		24 : {
			f: Qf
		},
		25 : {
			f: us
		},
		26 : {},
		27 : {},
		28 : {
			f: fl
		},
		29 : {},
		34 : {
			f: us
		},
		35 : {
			f: qf
		},
		38 : {
			f: An
		},
		39 : {
			f: An
		},
		40 : {
			f: An
		},
		41 : {
			f: An
		},
		42 : {
			f: us
		},
		43 : {
			f: us
		},
		47 : {
			f: Bo
		},
		49 : {
			f: wf
		},
		51 : {
			f: ds
		},
		60 : {},
		61 : {
			f: vf
		},
		64 : {
			f: us
		},
		65 : {
			f: gf
		},
		66 : {
			f: ds
		},
		77 : {},
		80 : {},
		81 : {},
		82 : {},
		85 : {
			f: ds
		},
		89 : {},
		90 : {},
		91 : {},
		92 : {
			f: rf
		},
		93 : {
			f: ul
		},
		94 : {},
		95 : {
			f: us
		},
		96 : {},
		97 : {},
		99 : {
			f: us
		},
		125 : {
			f: _l
		},
		128 : {
			f: $f
		},
		129 : {
			f: af
		},
		130 : {
			f: ds
		},
		131 : {
			f: us
		},
		132 : {
			f: us
		},
		133 : {
			f: nf
		},
		134 : {},
		140 : {
			f: wl
		},
		141 : {
			f: ds
		},
		144 : {},
		146 : {
			f: yl
		},
		151 : {},
		152 : {},
		153 : {},
		154 : {},
		155 : {},
		156 : {
			f: ds
		},
		157 : {},
		158 : {},
		160 : {
			f: Ol
		},
		161 : {
			f: xl
		},
		174 : {},
		175 : {},
		176 : {},
		177 : {},
		178 : {},
		180 : {},
		181 : {},
		182 : {},
		184 : {},
		185 : {},
		189 : {
			f: Ff
		},
		190 : {
			f: Df
		},
		193 : {
			f: ls
		},
		197 : {},
		198 : {},
		199 : {},
		200 : {},
		201 : {},
		202 : {
			f: us
		},
		203 : {},
		204 : {},
		205 : {},
		206 : {},
		207 : {},
		208 : {},
		209 : {},
		210 : {},
		211 : {},
		213 : {},
		215 : {},
		216 : {},
		217 : {},
		218 : {
			f: ds
		},
		220 : {},
		221 : {
			f: us
		},
		222 : {},
		224 : {
			f: Lf
		},
		225 : {
			f: ef
		},
		226 : {
			f: ls
		},
		227 : {},
		229 : {
			f: ol
		},
		233 : {},
		235 : {},
		236 : {},
		237 : {},
		239 : {},
		240 : {},
		241 : {},
		242 : {},
		244 : {},
		245 : {},
		246 : {},
		247 : {},
		248 : {},
		249 : {},
		251 : {},
		252 : {
			f: ff
		},
		253 : {
			f: Tf
		},
		255 : {
			f: of
		},
		256 : {},
		259 : {},
		290 : {},
		311 : {},
		312 : {},
		315 : {},
		317 : {
			f: ps
		},
		318 : {},
		319 : {},
		320 : {},
		330 : {},
		331 : {},
		333 : {},
		334 : {},
		335 : {},
		336 : {},
		337 : {},
		338 : {},
		339 : {},
		340 : {},
		351 : {},
		352 : {
			f: us
		},
		353 : {
			f: ls
		},
		401 : {},
		402 : {},
		403 : {},
		404 : {},
		405 : {},
		406 : {},
		407 : {},
		408 : {},
		425 : {},
		426 : {},
		427 : {},
		428 : {},
		429 : {},
		430 : {
			f: Jf
		},
		431 : {
			f: us
		},
		432 : {},
		433 : {},
		434 : {},
		437 : {},
		438 : {
			f: vl
		},
		439 : {
			f: us
		},
		440 : {
			f: pl
		},
		441 : {},
		442 : {
			f: ys
		},
		443 : {},
		444 : {
			f: ds
		},
		445 : {},
		446 : {},
		448 : {
			f: ls
		},
		449 : {
			f: hf,
			r: 2
		},
		450 : {
			f: ls
		},
		512 : {
			f: Of
		},
		513 : {
			f: Rl
		},
		515 : {
			f: Kf
		},
		516 : {
			f: Ef
		},
		517 : {
			f: Gf
		},
		519 : {
			f: Il
		},
		520 : {
			f: cf
		},
		523 : {},
		545 : {
			f: nl
		},
		549 : {
			f: df
		},
		566 : {},
		574 : {
			f: mf
		},
		638 : {
			f: Nf
		},
		659 : {},
		1048 : {},
		1054 : {
			f: Sf
		},
		1084 : {},
		1212 : {
			f: al
		},
		2048 : {
			f: bl
		},
		2049 : {},
		2050 : {},
		2051 : {},
		2052 : {},
		2053 : {},
		2054 : {},
		2055 : {},
		2056 : {},
		2057 : {
			f: qs
		},
		2058 : {},
		2059 : {},
		2060 : {},
		2061 : {},
		2062 : {},
		2063 : {},
		2064 : {},
		2066 : {},
		2067 : {},
		2128 : {},
		2129 : {},
		2130 : {},
		2131 : {},
		2132 : {},
		2133 : {},
		2134 : {},
		2135 : {},
		2136 : {},
		2137 : {},
		2138 : {},
		2146 : {},
		2147 : {
			r: 12
		},
		2148 : {},
		2149 : {},
		2150 : {},
		2151 : {
			f: ls
		},
		2152 : {},
		2154 : {},
		2155 : {},
		2156 : {},
		2161 : {},
		2162 : {},
		2164 : {},
		2165 : {},
		2166 : {},
		2167 : {},
		2168 : {},
		2169 : {},
		2170 : {},
		2171 : {},
		2172 : {
			f: El,
			r: 12
		},
		2173 : {
			f: uu,
			r: 12
		},
		2174 : {},
		2175 : {},
		2180 : {},
		2181 : {},
		2182 : {},
		2183 : {},
		2184 : {},
		2185 : {},
		2186 : {},
		2187 : {},
		2188 : {
			f: us,
			r: 12
		},
		2189 : {},
		2190 : {
			r: 12
		},
		2191 : {},
		2192 : {},
		2194 : {},
		2195 : {},
		2196 : {
			f: tl,
			r: 12
		},
		2197 : {},
		2198 : {
			f: iu,
			r: 12
		},
		2199 : {},
		2200 : {},
		2201 : {},
		2202 : {
			f: il,
			r: 12
		},
		2203 : {
			f: ls
		},
		2204 : {},
		2205 : {},
		2206 : {},
		2207 : {},
		2211 : {
			f: uf
		},
		2212 : {},
		2213 : {},
		2214 : {},
		2215 : {},
		4097 : {},
		4098 : {},
		4099 : {},
		4102 : {},
		4103 : {},
		4105 : {},
		4106 : {},
		4107 : {},
		4108 : {},
		4109 : {},
		4116 : {},
		4117 : {},
		4118 : {},
		4119 : {},
		4120 : {},
		4121 : {},
		4122 : {},
		4123 : {},
		4124 : {},
		4125 : {},
		4126 : {},
		4127 : {},
		4128 : {},
		4129 : {},
		4130 : {},
		4132 : {},
		4133 : {},
		4134 : {
			f: ds
		},
		4135 : {},
		4146 : {},
		4147 : {},
		4148 : {},
		4149 : {},
		4154 : {},
		4156 : {},
		4157 : {},
		4158 : {},
		4159 : {},
		4160 : {},
		4161 : {},
		4163 : {},
		4164 : {
			f: Al
		},
		4165 : {},
		4166 : {},
		4168 : {},
		4170 : {},
		4171 : {},
		4174 : {},
		4175 : {},
		4176 : {},
		4177 : {},
		4187 : {},
		4188 : {
			f: Tl
		},
		4189 : {},
		4191 : {},
		4192 : {},
		4193 : {},
		4194 : {},
		4195 : {},
		4196 : {},
		4197 : {},
		4198 : {},
		4199 : {},
		4200 : {},
		0 : {
			f: Of
		},
		1 : {},
		2 : {
			f: Ml
		},
		3 : {
			f: Pl
		},
		4 : {
			f: Dl
		},
		5 : {
			f: Wl
		},
		7 : {
			f: Bl
		},
		8 : {},
		9 : {
			f: qs
		},
		11 : {},
		22 : {
			f: ds
		},
		30 : {
			f: Af
		},
		31 : {},
		32 : {},
		33 : {
			f: nl
		},
		36 : {},
		37 : {
			f: df
		},
		50 : {
			f: zl
		},
		62 : {},
		52 : {},
		67 : {
			f: Uf
		},
		68 : {
			f: ds
		},
		69 : {},
		86 : {},
		126 : {},
		127 : {
			f: Nl
		},
		135 : {},
		136 : {},
		137 : {},
		143 : {
			f: Vl
		},
		145 : {},
		148 : {},
		149 : {},
		150 : {},
		169 : {},
		171 : {},
		188 : {},
		191 : {},
		192 : {},
		194 : {},
		195 : {},
		214 : {
			f: Hl
		},
		223 : {},
		234 : {},
		354 : {},
		421 : {},
		518 : {
			f: $d
		},
		521 : {
			f: qs
		},
		536 : {
			f: Qf
		},
		547 : {
			f: qf
		},
		561 : {},
		579 : {
			f: Hf
		},
		1030 : {
			f: $d
		},
		1033 : {
			f: qs
		},
		1091 : {
			f: Vf
		},
		2157 : {},
		2163 : {},
		2177 : {},
		2240 : {},
		2241 : {},
		2242 : {},
		2243 : {},
		2244 : {},
		2245 : {},
		2246 : {},
		2247 : {},
		2248 : {},
		2249 : {},
		2250 : {},
		2251 : {},
		2262 : {
			r: 12
		},
		101 : {},
		102 : {},
		105 : {},
		106 : {},
		107 : {},
		109 : {},
		112 : {},
		114 : {},
		29282 : {}
	};
	function rg(e, r, t, a) {
		var n = r;
		if (isNaN(n)) return;
		var i = a || (t || []).length || 0;
		var s = e.next(4);
		s._W(2, n);
		s._W(2, i);
		if (i > 0 && fa(t)) e.push(t)
	}
	function tg(e, r, t, a) {
		var n = a || (t || []).length || 0;
		if (n <= 8224) return rg(e, r, t, n);
		var i = r;
		if (isNaN(i)) return;
		var s = t.parts || [],
		f = 0;
		var l = 0,
		o = 0;
		while (o + (s[f] || 8224) <= 8224) {
			o += s[f] || 8224;
			f++
		}
		var c = e.next(4);
		c._W(2, i);
		c._W(2, o);
		e.push(t.slice(l, l + o));
		l += o;
		while (l < n) {
			c = e.next(4);
			c._W(2, 60);
			o = 0;
			while (o + (s[f] || 8224) <= 8224) {
				o += s[f] || 8224;
				f++
			}
			c._W(2, o);
			e.push(t.slice(l, l + o));
			l += o
		}
	}
	function ag(e, r, t, a) {
		var n = Ea(9);
		Fl(n, e, r);
		bs(t, a || "b", n);
		return n
	}
	function ng(e, r, t) {
		var a = Ea(8 + 2 * t.length);
		Fl(a, e, r);
		a._W(1, t.length);
		a._W(t.length, t, "sbcs");
		return a.l < a.length ? a.slice(0, a.l) : a
	}
	function ig(e, r) {
		r.forEach(function(r) {
			var t = r[0].map(function(e) {
				return e.t
			}).join("");
			if (t.length <= 2048) return rg(e, 28, ll(t, r[1], r[2]));
			rg(e, 28, ll(t.slice(0, 2048), r[1], r[2], t.length));
			for (var a = 2048; a < t.length; a += 2048) rg(e, 28, ll(t.slice(a, Math.min(a + 2048, t.length)), -1, -1, Math.min(2048, t.length - a)))
		})
	}
	function sg(e, r, t, a, n, i) {
		var s = 0;
		if (r.z != null) {
			s = n._BIFF2FmtTable.indexOf(r.z);
			if (s == -1) {
				n._BIFF2FmtTable.push(r.z);
				s = n._BIFF2FmtTable.length - 1
			}
		}
		var f = 0;
		if (r.z != null) {
			for (; f < n.cellXfs.length; ++f) if (n.cellXfs[f].numFmtId == s) break;
			if (f == n.cellXfs.length) n.cellXfs.push({
				numFmtId: s
			})
		}
		if (r.v != null) switch (r.t) {
		case "d":
			;
		case "n":
			var l = r.t == "d" ? dr(wr(r.v, i), i) : r.v;
			if (n.biff == 2 && l == (l | 0) && l >= 0 && l < 65536) rg(e, 2, Ul(t, a, l, f, s));
			else if (isNaN(l)) rg(e, 5, ag(t, a, 36, "e"));
			else if (!isFinite(l)) rg(e, 5, ag(t, a, 7, "e"));
			else rg(e, 3, Ll(t, a, l, f, s));
			return;
		case "b":
			;
		case "e":
			rg(e, 5, ag(t, a, r.v, r.t));
			return;
		case "s":
			;
		case "str":
			rg(e, 4, ng(t, a, r.v == null ? "": String(r.v).slice(0, 255)));
			return;
		}
		rg(e, 1, Fl(null, t, a))
	}
	function fg(e, r, t, a, n) {
		var i = r["!data"] != null;
		var s = Ga(r["!ref"] || "A1"),
		f,
		l = "",
		o = [];
		if (s.e.c > 255 || s.e.r > 16383) {
			if (a.WTF) throw new Error("Range " + (r["!ref"] || "A1") + " exceeds format limit A1:IV16384");
			s.e.c = Math.min(s.e.c, 255);
			s.e.r = Math.min(s.e.c, 16383)
		}
		var c = (((n || {}).Workbook || {}).WBProps || {}).date1904;
		var u = [],
		h = [];
		for (var d = s.s.c; d <= s.e.c; ++d) o[d] = La(d);
		for (var v = s.s.r; v <= s.e.r; ++v) {
			if (i) u = r["!data"][v] || [];
			l = Na(v);
			for (d = s.s.c; d <= s.e.c; ++d) {
				var p = i ? u[d] : r[o[d] + l];
				if (!p) continue;
				sg(e, p, v, d, a, c);
				if (p.c) h.push([p.c, v, d])
			}
		}
		ig(e, h)
	}
	function lg(e, r) {
		var t = r || {};
		var a = Sa();
		var n = 0;
		for (var i = 0; i < e.SheetNames.length; ++i) if (e.SheetNames[i] == t.sheet) n = i;
		if (n == 0 && !!t.sheet && e.SheetNames[0] != t.sheet) throw new Error("Sheet not found: " + t.sheet);
		rg(a, t.biff == 4 ? 1033 : t.biff == 3 ? 521 : 9, Qs(e, 16, t));
		if (((e.Workbook || {}).WBProps || {}).date1904) rg(a, 34, hs(true));
		t.cellXfs = [{
			numFmtId: 0
		}];
		t._BIFF2FmtTable = ["General"];
		t._Fonts = [];
		var s = Sa();
		fg(s, e.Sheets[e.SheetNames[n]], n, t, e);
		t._BIFF2FmtTable.forEach(function(e) {
			if (t.biff <= 3) rg(a, 30, Cf(e));
			else rg(a, 1054, Rf(e))
		});
		t.cellXfs.forEach(function(e) {
			switch (t.biff) {
			case 2:
				rg(a, 67, Bf(e));
				break;
			case 3:
				rg(a, 579, Wf(e));
				break;
			case 4:
				rg(a, 1091, zf(e));
				break;
			}
		});
		delete t._BIFF2FmtTable;
		delete t.cellXfs;
		delete t._Fonts;
		a.push(s.end());
		rg(a, 10);
		return a.end()
	}
	var og = 1,
	cg = [];
	function ug() {
		var e = Ea(82 + 8 * cg.length);
		e._W(2, 15);
		e._W(2, 61440);
		e._W(4, 74 + 8 * cg.length); {
			e._W(2, 0);
			e._W(2, 61446);
			e._W(4, 16 + 8 * cg.length); {
				e._W(4, og);
				e._W(4, cg.length + 1);
				var r = 0;
				for (var t = 0; t < cg.length; ++t) r += cg[t] && cg[t][1] || 0;
				e._W(4, r);
				e._W(4, cg.length)
			}
			cg.forEach(function(r) {
				e._W(4, r[0]);
				e._W(4, r[2])
			})
		} {
			e._W(2, 51);
			e._W(2, 61451);
			e._W(4, 18);
			e._W(2, 191);
			e._W(4, 524296);
			e._W(2, 385);
			e._W(4, 134217793);
			e._W(2, 448);
			e._W(4, 134217792)
		} {
			e._W(2, 64);
			e._W(2, 61726);
			e._W(4, 16);
			e._W(4, 134217741);
			e._W(4, 134217740);
			e._W(4, 134217751);
			e._W(4, 268435703)
		}
		return e
	}
	function hg(e, r) {
		var t = [],
		a = 0,
		n = Sa(),
		i = og;
		var s;
		r.forEach(function(e, r) {
			var i = "";
			var f = e[0].map(function(e) {
				if (e.a && !i) i = e.a;
				return e.t
			}).join(""); ++og; {
				var l = Ea(150);
				l._W(2, 15);
				l._W(2, 61444);
				l._W(4, 150); {
					l._W(2, 3234);
					l._W(2, 61450);
					l._W(4, 8);
					l._W(4, og);
					l._W(4, 2560)
				} {
					l._W(2, 227);
					l._W(2, 61451);
					l._W(4, 84);
					l._W(2, 128);
					l._W(4, 0);
					l._W(2, 139);
					l._W(4, 2);
					l._W(2, 191);
					l._W(4, 524296);
					l._W(2, 344);
					l.l += 4;
					l._W(2, 385);
					l._W(4, 134217808);
					l._W(2, 387);
					l._W(4, 134217808);
					l._W(2, 389);
					l._W(4, 268435700);
					l._W(2, 447);
					l._W(4, 1048592);
					l._W(2, 448);
					l._W(4, 134217809);
					l._W(2, 451);
					l._W(4, 268435700);
					l._W(2, 513);
					l._W(4, 134217809);
					l._W(2, 515);
					l._W(4, 268435700);
					l._W(2, 575);
					l._W(4, 196609);
					l._W(2, 959);
					l._W(4, 131072 | (e[0].hidden ? 2 : 0))
				} {
					l.l += 2;
					l._W(2, 61456);
					l._W(4, 18);
					l._W(2, 3);
					l._W(2, e[2] + 2);
					l.l += 2;
					l._W(2, e[1] + 1);
					l.l += 2;
					l._W(2, e[2] + 4);
					l.l += 2;
					l._W(2, e[1] + 5);
					l.l += 2
				} {
					l.l += 2;
					l._W(2, 61457);
					l.l += 4
				}
				l.l = 150;
				if (r == 0) s = l;
				else rg(n, 236, l)
			}
			a += 150; {
				var o = Ea(52);
				o._W(2, 21);
				o._W(2, 18);
				o._W(2, 25);
				o._W(2, og);
				o._W(2, 0);
				o.l = 22;
				o._W(2, 13);
				o._W(2, 22);
				o._W(4, 1651663474);
				o._W(4, 2503426821);
				o._W(4, 2150634280);
				o._W(4, 1768515844 + og * 256);
				o._W(2, 0);
				o._W(4, 0);
				o.l += 4;
				rg(n, 93, o)
			} {
				var c = Ea(8);
				c.l += 2;
				c._W(2, 61453);
				c.l += 4;
				rg(n, 236, c)
			}
			a += 8; {
				var u = Ea(18);
				u._W(2, 18);
				u.l += 8;
				u._W(2, f.length);
				u._W(2, 16);
				u.l += 4;
				rg(n, 438, u); {
					var h = Ea(1 + f.length);
					h._W(1, 0);
					h._W(f.length, f, "sbcs");
					rg(n, 60, h)
				} {
					var d = Ea(16);
					d.l += 8;
					d._W(2, f.length);
					d.l += 6;
					rg(n, 60, d)
				}
			} {
				var v = Ea(12 + i.length);
				v._W(2, e[1]);
				v._W(2, e[2]);
				v._W(2, 0 | (e[0].hidden ? 0 : 2));
				v._W(2, og);
				v._W(2, i.length);
				v._W(1, 0);
				v._W(i.length, i, "sbcs");
				v.l++;
				t.push(v)
			}
		}); {
			var f = Ea(80);
			f._W(2, 15);
			f._W(2, 61442);
			f._W(4, a + f.length - 8); {
				f._W(2, 16);
				f._W(2, 61448);
				f._W(4, 8);
				f._W(4, r.length + 1);
				f._W(4, og)
			} {
				f._W(2, 15);
				f._W(2, 61443);
				f._W(4, a + 48); {
					f._W(2, 15);
					f._W(2, 61444);
					f._W(4, 40); {
						f._W(2, 1);
						f._W(2, 61449);
						f._W(4, 16);
						f.l += 16
					} {
						f._W(2, 2);
						f._W(2, 61450);
						f._W(4, 8);
						f._W(4, i);
						f._W(4, 5)
					}
				}
			}
			rg(e, 236, s ? P([f, s]) : f)
		}
		e.push(n.end());
		t.forEach(function(r) {
			rg(e, 28, r)
		});
		cg.push([i, r.length + 1, og]); ++og
	}
	function dg(e, r, t) {
		rg(e, 49, kf({
			sz: 12,
			color: {
				theme: 1
			},
			name: "Arial",
			family: 2,
			scheme: "minor"
		},
		t))
	}
	function vg(e, r, t) {
		if (!r) return; [[5, 8], [23, 26], [41, 44], [50, 392]].forEach(function(a) {
			for (var n = a[0]; n <= a[1]; ++n) if (r[n] != null) rg(e, 1054, xf(n, r[n], t))
		})
	}
	function pg(e, r) {
		var t = Ea(19);
		t._W(4, 2151);
		t._W(4, 0);
		t._W(4, 0);
		t._W(2, 3);
		t._W(1, 1);
		t._W(4, 0);
		rg(e, 2151, t);
		t = Ea(39);
		t._W(4, 2152);
		t._W(4, 0);
		t._W(4, 0);
		t._W(2, 3);
		t._W(1, 0);
		t._W(4, 0);
		t._W(2, 1);
		t._W(4, 4);
		t._W(2, 0);
		Vs(Ga(r["!ref"] || "A1"), t);
		t._W(4, 4);
		rg(e, 2152, t)
	}
	function mg(e, r) {
		for (var t = 0; t < 16; ++t) rg(e, 224, Mf({
			numFmtId: 0,
			style: true
		},
		0, r));
		r.cellXfs.forEach(function(t) {
			rg(e, 224, Mf(t, 0, r))
		})
	}
	function bg(e, r) {
		for (var t = 0; t < r["!links"].length; ++t) {
			var a = r["!links"][t];
			rg(e, 440, ml(a));
			if (a[1].Tooltip) rg(e, 2048, gl(a))
		}
		delete r["!links"]
	}
	function gg(e, r) {
		if (!r) return;
		var t = 0;
		r.forEach(function(r, a) {
			if (++t <= 256 && r) {
				rg(e, 125, Sl(kv(a, r), a))
			}
		})
	}
	function wg(e, r, t, a, n, i) {
		var s = 16 + yv(n.cellXfs, r, n);
		if (r.v == null && !r.bf) {
			rg(e, 513, Ls(t, a, s));
			return
		}
		if (r.bf) rg(e, 6, Xd(r, t, a, n, s));
		else switch (r.t) {
		case "d":
			;
		case "n":
			var f = r.t == "d" ? dr(wr(r.v, i), i) : r.v;
			if (isNaN(f)) rg(e, 517, jf(t, a, 36, s, n, "e"));
			else if (!isFinite(f)) rg(e, 517, jf(t, a, 7, s, n, "e"));
			else rg(e, 515, Yf(t, a, f, s, n));
			break;
		case "b":
			;
		case "e":
			rg(e, 517, jf(t, a, r.v, s, n, r.t));
			break;
		case "s":
			;
		case "str":
			if (n.bookSST) {
				var l = wv(n.Strings, r.v == null ? "": String(r.v), n.revStrings);
				rg(e, 253, yf(t, a, l, s, n))
			} else rg(e, 516, _f(t, a, (r.v == null ? "": String(r.v)).slice(0, 255), s, n));
			break;
		default:
			rg(e, 513, Ls(t, a, s));
		}
	}
	function kg(e, r, t) {
		var a = Sa();
		var n = t.SheetNames[e],
		i = t.Sheets[n] || {};
		var s = (t || {}).Workbook || {};
		var f = (s.Sheets || [])[e] || {};
		var l = i["!data"] != null;
		var o = r.biff == 8;
		var c, u = "",
		h = [];
		var d = Ga(i["!ref"] || "A1");
		var v = o ? 65536 : 16384;
		if (d.e.c > 255 || d.e.r >= v) {
			if (r.WTF) throw new Error("Range " + (i["!ref"] || "A1") + " exceeds format limit A1:IV16384");
			d.e.c = Math.min(d.e.c, 255);
			d.e.r = Math.min(d.e.c, v - 1)
		}
		rg(a, 2057, Qs(t, 16, r));
		rg(a, 13, vs(1));
		rg(a, 12, vs(100));
		rg(a, 15, hs(true));
		rg(a, 17, hs(false));
		rg(a, 16, Cn(.001));
		rg(a, 95, hs(true));
		rg(a, 42, hs(false));
		rg(a, 43, hs(false));
		rg(a, 130, vs(1));
		rg(a, 128, Xf([0, 0]));
		rg(a, 131, hs(false));
		rg(a, 132, hs(false));
		if (o) gg(a, i["!cols"]);
		rg(a, 512, If(d, r));
		var p = (((t || {}).Workbook || {}).WBProps || {}).date1904;
		if (o) i["!links"] = [];
		var m = [];
		var b = [];
		for (var g = d.s.c; g <= d.e.c; ++g) h[g] = La(g);
		for (var w = d.s.r; w <= d.e.r; ++w) {
			if (l) b = i["!data"][w] || [];
			u = Na(w);
			for (g = d.s.c; g <= d.e.c; ++g) {
				c = h[g] + u;
				var k = l ? b[g] : i[c];
				if (!k) continue;
				wg(a, k, w, g, r, p);
				if (o && k.l) i["!links"].push([c, k.l]);
				if (k.c) m.push([k.c, w, g])
			}
		}
		var T = f.CodeName || f.name || n;
		if (o) hg(a, m);
		else ig(a, m);
		if (o) rg(a, 574, bf((s.Views || [])[0]));
		if (o && (i["!merges"] || []).length) rg(a, 229, cl(i["!merges"]));
		if (o) bg(a, i);
		rg(a, 442, _s(T, r));
		if (o) pg(a, i);
		rg(a, 10);
		return a.end()
	}
	function Tg(e, r, t) {
		var a = Sa();
		var n = (e || {}).Workbook || {};
		var i = n.Sheets || [];
		var s = n.WBProps || {};
		var f = t.biff == 8,
		l = t.biff == 5;
		rg(a, 2057, Qs(e, 5, t));
		if (t.bookType == "xla") rg(a, 135);
		rg(a, 225, f ? vs(1200) : null);
		rg(a, 193, os(2));
		if (l) rg(a, 191);
		if (l) rg(a, 192);
		rg(a, 226);
		rg(a, 92, tf("SheetJS", t));
		rg(a, 66, vs(f ? 1200 : 1252));
		if (f) rg(a, 353, vs(0));
		if (f) rg(a, 448);
		rg(a, 317, Cl(e.SheetNames.length));
		if (f && e.vbaraw) rg(a, 211);
		if (f && e.vbaraw) {
			var o = s.CodeName || "ThisWorkbook";
			rg(a, 442, _s(o, t))
		}
		rg(a, 156, vs(17));
		rg(a, 25, hs(false));
		rg(a, 18, hs(false));
		rg(a, 19, vs(0));
		if (f) rg(a, 431, hs(false));
		if (f) rg(a, 444, vs(0));
		rg(a, 61, pf(t));
		rg(a, 64, hs(false));
		rg(a, 141, vs(0));
		rg(a, 34, hs(Dm(e) == "true"));
		rg(a, 14, hs(true));
		if (f) rg(a, 439, hs(false));
		rg(a, 218, vs(0));
		dg(a, e, t);
		vg(a, e.SSF, t);
		mg(a, t);
		if (f) rg(a, 352, hs(false));
		var c = a.end();
		var u = Sa();
		if (f) rg(u, 140, kl());
		if (f && cg.length) rg(u, 235, ug());
		if (f && t.Strings) tg(u, 252, lf(t.Strings, t));
		rg(u, 10);
		var h = u.end();
		var d = Sa();
		var v = 0,
		p = 0;
		for (p = 0; p < e.SheetNames.length; ++p) v += (f ? 12 : 11) + (f ? 2 : 1) * e.SheetNames[p].length;
		var m = c.length + v + h.length;
		for (p = 0; p < e.SheetNames.length; ++p) {
			var b = i[p] || {};
			rg(d, 133, sf({
				pos: m,
				hs: b.Hidden || 0,
				dt: 0,
				name: e.SheetNames[p]
			},
			t));
			m += r[p].length
		}
		var g = d.end();
		if (v != g.length) throw new Error("BS8 " + v + " != " + g.length);
		var w = [];
		if (c.length) w.push(c);
		if (g.length) w.push(g);
		if (h.length) w.push(h);
		return P(w)
	}
	function yg(e, r) {
		var t = r || {};
		var a = [];
		if (e && !e.SSF) {
			e.SSF = Tr(q)
		}
		if (e && e.SSF) {
			$e();
			Ve(e.SSF);
			t.revssf = lr(e.SSF);
			t.revssf[e.SSF[65535]] = 0;
			t.ssf = e.SSF
		}
		og = 1;
		cg = [];
		t.Strings = [];
		t.Strings.Count = 0;
		t.Strings.Unique = 0;
		rk(t);
		t.cellXfs = [];
		yv(t.cellXfs, {},
		{
			revssf: {
				General: 0
			}
		});
		if (!e.Props) e.Props = {};
		for (var n = 0; n < e.SheetNames.length; ++n) a[a.length] = kg(n, t, e);
		a.unshift(Tg(e, a, t));
		return P(a)
	}
	function Eg(e, r) {
		for (var t = 0; t <= e.SheetNames.length; ++t) {
			var a = e.Sheets[e.SheetNames[t]];
			if (!a || !a["!ref"]) continue;
			var n = Ha(a["!ref"]);
			if (n.e.c > 255) {
				if (typeof console != "undefined" && console.error) console.error("Worksheet '" + e.SheetNames[t] + "' extends beyond column IV (255).  Data may be lost.")
			}
		}
		var i = r || {};
		switch (i.biff || 2) {
		case 8:
			;
		case 5:
			return yg(e, r);
		case 4:
			;
		case 3:
			;
		case 2:
			return lg(e, r);
		}
		throw new Error("invalid type " + i.bookType + " for BIFF")
	}
	function _g(e, r) {
		var t = r || {};
		var a = t.dense != null ? t.dense: g;
		var n = {};
		if (a) n["!data"] = [];
		e = e.replace(/<!--.*?-->/g, "");
		var i = e.match(/<table/i);
		if (!i) throw new Error("Invalid HTML: could not find <table>");
		var s = e.match(/<\/table/i);
		var f = i.index,
		l = s && s.index || e.length;
		var o = Nr(e.slice(f, l), /(:?<tr[^>]*>)/i, "<tr>");
		var c = -1,
		u = 0,
		h = 0,
		d = 0;
		var v = {
			s: {
				r: 1e7,
				c: 1e7
			},
			e: {
				r: 0,
				c: 0
			}
		};
		var p = [];
		for (f = 0; f < o.length; ++f) {
			var m = o[f].trim();
			var b = m.slice(0, 3).toLowerCase();
			if (b == "<tr") {++c;
				if (t.sheetRows && t.sheetRows <= c) {--c;
					break
				}
				u = 0;
				continue
			}
			if (b != "<td" && b != "<th") continue;
			var w = m.split(/<\/t[dh]>/i);
			for (l = 0; l < w.length; ++l) {
				var k = w[l].trim();
				if (!k.match(/<t[dh]/i)) continue;
				var T = k,
				y = 0;
				while (T.charAt(0) == "<" && (y = T.indexOf(">")) > -1) T = T.slice(y + 1);
				for (var E = 0; E < p.length; ++E) {
					var _ = p[E];
					if (_.s.c == u && _.s.r < c && c <= _.e.r) {
						u = _.e.c + 1;
						E = -1
					}
				}
				var S = rt(k.slice(0, k.indexOf(">")));
				d = S.colspan ? +S.colspan: 1;
				if ((h = +S.rowspan) > 1 || d > 1) p.push({
					s: {
						r: c,
						c: u
					},
					e: {
						r: c + (h || 1) - 1,
						c: u + d - 1
					}
				});
				var x = S.t || S["data-t"] || "";
				if (!T.length) {
					u += d;
					continue
				}
				T = Et(T);
				if (v.s.r > c) v.s.r = c;
				if (v.e.r < c) v.e.r = c;
				if (v.s.c > u) v.s.c = u;
				if (v.e.c < u) v.e.c = u;
				if (!T.length) {
					u += d;
					continue
				}
				var A = {
					t: "s",
					v: T
				};
				if (t.raw || !T.trim().length || x == "s") {} else if (T === "TRUE") A = {
					t: "b",
					v: true
				};
				else if (T === "FALSE") A = {
					t: "b",
					v: false
				};
				else if (!isNaN(Er(T))) A = {
					t: "n",
					v: Er(T)
				};
				else if (!isNaN(Ir(T).getDate())) {
					A = {
						t: "d",
						v: wr(T)
					};
					if (t.UTC === false) A.v = Fr(A.v);
					if (!t.cellDates) A = {
						t: "n",
						v: dr(A.v)
					};
					A.z = t.dateNF || q[14]
				}
				if (A.cellText !== false) A.w = T;
				if (a) {
					if (!n["!data"][c]) n["!data"][c] = [];
					n["!data"][c][u] = A
				} else n[za({
					r: c,
					c: u
				})] = A;
				u += d
			}
		}
		n["!ref"] = Va(v);
		if (p.length) n["!merges"] = p;
		return n
	}
	function Sg(e, r, t, a) {
		var n = e["!merges"] || [];
		var i = [];
		var s = {};
		var f = e["!data"] != null;
		for (var l = r.s.c; l <= r.e.c; ++l) {
			var o = 0,
			c = 0;
			for (var u = 0; u < n.length; ++u) {
				if (n[u].s.r > t || n[u].s.c > l) continue;
				if (n[u].e.r < t || n[u].e.c < l) continue;
				if (n[u].s.r < t || n[u].s.c < l) {
					o = -1;
					break
				}
				o = n[u].e.r - n[u].s.r + 1;
				c = n[u].e.c - n[u].s.c + 1;
				break
			}
			if (o < 0) continue;
			var h = La(l) + Na(t);
			var d = f ? (e["!data"][t] || [])[l] : e[h];
			var v = d && d.v != null && (d.h || ut(d.w || (Ka(d), d.w) || "")) || "";
			s = {};
			if (o > 1) s.rowspan = o;
			if (c > 1) s.colspan = c;
			if (a.editable) v = '<span contenteditable="true">' + v + "</span>";
			else if (d) {
				s["data-t"] = d && d.t || "z";
				if (d.v != null) s["data-v"] = d.v instanceof Date ? d.v.toISOString() : d.v;
				if (d.z != null) s["data-z"] = d.z;
				if (d.l && (d.l.Target || "#").charAt(0) != "#") v = '<a href="' + ut(d.l.Target) + '">' + v + "</a>"
			}
			s.id = (a.id || "sjs") + "-" + h;
			i.push(It("td", v, s))
		}
		var p = "<tr>";
		return p + i.join("") + "</tr>"
	}
	var xg = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
	var Ag = "</body></html>";
	function Cg(e, r) {
		var t = e.match(/<table[\s\S]*?>[\s\S]*?<\/table>/gi);
		if (!t || t.length == 0) throw new Error("Invalid HTML: could not find <table>");
		if (t.length == 1) {
			var a = Ya(_g(t[0], r), r);
			a.bookType = "html";
			return a
		}
		var n = jk();
		t.forEach(function(e, t) {
			Kk(n, _g(e, r), "Sheet" + (t + 1))
		});
		n.bookType = "html";
		return n
	}
	function Rg(e, r, t) {
		var a = [];
		return a.join("") + "<table" + (t && t.id ? ' id="' + t.id + '"': "") + ">"
	}
	function Og(e, r) {
		var t = r || {};
		var a = t.header != null ? t.header: xg;
		var n = t.footer != null ? t.footer: Ag;
		var i = [a];
		var s = Ha(e["!ref"] || "A1");
		i.push(Rg(e, s, t));
		if (e["!ref"]) for (var f = s.s.r; f <= s.e.r; ++f) i.push(Sg(e, s, f, t));
		i.push("</table>" + n);
		return i.join("")
	}
	function Ig(e, r, t) {
		var a = r.rows;
		if (!a) {
			throw "Unsupported origin when " + r.tagName + " is not a TABLE"
		}
		var n = t || {};
		var i = e["!data"] != null;
		var s = 0,
		f = 0;
		if (n.origin != null) {
			if (typeof n.origin == "number") s = n.origin;
			else {
				var l = typeof n.origin == "string" ? Wa(n.origin) : n.origin;
				s = l.r;
				f = l.c
			}
		}
		var o = Math.min(n.sheetRows || 1e7, a.length);
		var c = {
			s: {
				r: 0,
				c: 0
			},
			e: {
				r: s,
				c: f
			}
		};
		if (e["!ref"]) {
			var u = Ha(e["!ref"]);
			c.s.r = Math.min(c.s.r, u.s.r);
			c.s.c = Math.min(c.s.c, u.s.c);
			c.e.r = Math.max(c.e.r, u.e.r);
			c.e.c = Math.max(c.e.c, u.e.c);
			if (s == -1) c.e.r = s = u.e.r + 1
		}
		var h = [],
		d = 0;
		var v = e["!rows"] || (e["!rows"] = []);
		var p = 0,
		m = 0,
		b = 0,
		g = 0,
		w = 0,
		k = 0;
		if (!e["!cols"]) e["!cols"] = [];
		for (; p < a.length && m < o; ++p) {
			var T = a[p];
			if (Dg(T)) {
				if (n.display) continue;
				v[m] = {
					hidden: true
				}
			}
			var y = T.cells;
			for (b = g = 0; b < y.length; ++b) {
				var E = y[b];
				if (n.display && Dg(E)) continue;
				var _ = E.hasAttribute("data-v") ? E.getAttribute("data-v") : E.hasAttribute("v") ? E.getAttribute("v") : Et(E.innerHTML);
				var S = E.getAttribute("data-z") || E.getAttribute("z");
				for (d = 0; d < h.length; ++d) {
					var x = h[d];
					if (x.s.c == g + f && x.s.r < m + s && m + s <= x.e.r) {
						g = x.e.c + 1 - f;
						d = -1
					}
				}
				k = +E.getAttribute("colspan") || 1;
				if ((w = +E.getAttribute("rowspan") || 1) > 1 || k > 1) h.push({
					s: {
						r: m + s,
						c: g + f
					},
					e: {
						r: m + s + (w || 1) - 1,
						c: g + f + (k || 1) - 1
					}
				});
				var A = {
					t: "s",
					v: _
				};
				var C = E.getAttribute("data-t") || E.getAttribute("t") || "";
				if (_ != null) {
					if (_.length == 0) A.t = C || "z";
					else if (n.raw || _.trim().length == 0 || C == "s") {} else if (_ === "TRUE") A = {
						t: "b",
						v: true
					};
					else if (_ === "FALSE") A = {
						t: "b",
						v: false
					};
					else if (!isNaN(Er(_))) A = {
						t: "n",
						v: Er(_)
					};
					else if (!isNaN(Ir(_).getDate())) {
						A = {
							t: "d",
							v: wr(_)
						};
						if (n.UTC) A.v = Dr(A.v);
						if (!n.cellDates) A = {
							t: "n",
							v: dr(A.v)
						};
						A.z = n.dateNF || q[14]
					}
				}
				if (A.z === undefined && S != null) A.z = S;
				var R = "",
				O = E.getElementsByTagName("A");
				if (O && O.length) for (var I = 0; I < O.length; ++I) if (O[I].hasAttribute("href")) {
					R = O[I].getAttribute("href");
					if (R.charAt(0) != "#") break
				}
				if (R && R.charAt(0) != "#" && R.slice(0, 11).toLowerCase() != "javascript:") A.l = {
					Target: R
				};
				if (i) {
					if (!e["!data"][m + s]) e["!data"][m + s] = [];
					e["!data"][m + s][g + f] = A
				} else e[za({
					c: g + f,
					r: m + s
				})] = A;
				if (c.e.c < g + f) c.e.c = g + f;
				g += k
			}++m
		}
		if (h.length) e["!merges"] = (e["!merges"] || []).concat(h);
		c.e.r = Math.max(c.e.r, m - 1 + s);
		e["!ref"] = Va(c);
		if (m >= o) e["!fullref"] = Va((c.e.r = a.length - p + m - 1 + s, c));
		return e
	}
	function Ng(e, r) {
		var t = r || {};
		var a = {};
		if (t.dense) a["!data"] = [];
		return Ig(a, e, r)
	}
	function Fg(e, r) {
		var t = Ya(Ng(e, r), r);
		return t
	}
	function Dg(e) {
		var r = "";
		var t = Pg(e);
		if (t) r = t(e).getPropertyValue("display");
		if (!r) r = e.style && e.style.display;
		return r === "none"
	}
	function Pg(e) {
		if (e.ownerDocument.defaultView && typeof e.ownerDocument.defaultView.getComputedStyle === "function") return e.ownerDocument.defaultView.getComputedStyle;
		if (typeof getComputedStyle === "function") return getComputedStyle;
		return null
	}
	function Lg(e) {
		var r = e.replace(/[\t\r\n]/g, " ").trim().replace(/ +/g, " ").replace(/<text:s\/>/g, " ").replace(/<text:s text:c="(\d+)"\/>/g,
		function(e, r) {
			return Array(parseInt(r, 10) + 1).join(" ")
		}).replace(/<text:tab[^>]*\/>/g, "\t").replace(/<text:line-break\/>/g, "\n");
		var t = it(r.replace(/<[^>]*>/g, ""));
		return [t]
	}
	function Mg(e, r, t) {
		var a = t || {};
		var n = Dt(e);
		Pt.lastIndex = 0;
		n = n.replace(/<!--([\s\S]*?)-->/gm, "").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm, "");
		var i, s, f = "",
		l = "",
		o, c = 0,
		u = -1,
		h = false,
		d = "";
		while (i = Pt.exec(n)) {
			switch (i[3] = i[3].replace(/_.*$/, "")) {
			case "number-style":
				;
			case "currency-style":
				;
			case "percentage-style":
				;
			case "date-style":
				;
			case "time-style":
				;
			case "text-style":
				if (i[1] === "/") {
					h = false;
					if (s["truncate-on-overflow"] == "false") {
						if (f.match(/h/)) f = f.replace(/h+/, "[$&]");
						else if (f.match(/m/)) f = f.replace(/m+/, "[$&]");
						else if (f.match(/s/)) f = f.replace(/s+/, "[$&]")
					}
					a[s.name] = f;
					f = ""
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					h = true;
					f = "";
					s = rt(i[0], false)
				}
				break;
			case "boolean-style":
				if (i[1] === "/") {
					h = false;
					a[s.name] = "General";
					f = ""
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					h = true;
					f = "";
					s = rt(i[0], false)
				}
				break;
			case "boolean":
				f += "General";
				break;
			case "text":
				if (i[1] === "/") {
					d = n.slice(u, Pt.lastIndex - i[0].length);
					if (d == "%" && s[0] == "<number:percentage-style") f += "%";
					else f += '"' + d.replace(/"/g, '""') + '"'
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					u = Pt.lastIndex
				}
				break;
			case "day":
				{
					o = rt(i[0], false);
					switch (o["style"]) {
					case "short":
						f += "d";
						break;
					case "long":
						f += "dd";
						break;
					default:
						f += "dd";
						break;
					}
				}
				break;
			case "day-of-week":
				{
					o = rt(i[0], false);
					switch (o["style"]) {
					case "short":
						f += "ddd";
						break;
					case "long":
						f += "dddd";
						break;
					default:
						f += "ddd";
						break;
					}
				}
				break;
			case "era":
				{
					o = rt(i[0], false);
					switch (o["style"]) {
					case "short":
						f += "ee";
						break;
					case "long":
						f += "eeee";
						break;
					default:
						f += "eeee";
						break;
					}
				}
				break;
			case "hours":
				{
					o = rt(i[0], false);
					switch (o["style"]) {
					case "short":
						f += "h";
						break;
					case "long":
						f += "hh";
						break;
					default:
						f += "hh";
						break;
					}
				}
				break;
			case "minutes":
				{
					o = rt(i[0], false);
					switch (o["style"]) {
					case "short":
						f += "m";
						break;
					case "long":
						f += "mm";
						break;
					default:
						f += "mm";
						break;
					}
				}
				break;
			case "month":
				{
					o = rt(i[0], false);
					if (o["textual"]) f += "mm";
					switch (o["style"]) {
					case "short":
						f += "m";
						break;
					case "long":
						f += "mm";
						break;
					default:
						f += "m";
						break;
					}
				}
				break;
			case "seconds":
				{
					o = rt(i[0], false);
					switch (o["style"]) {
					case "short":
						f += "s";
						break;
					case "long":
						f += "ss";
						break;
					default:
						f += "ss";
						break;
					}
					if (o["decimal-places"]) f += "." + yr("0", +o["decimal-places"])
				}
				break;
			case "year":
				{
					o = rt(i[0], false);
					switch (o["style"]) {
					case "short":
						f += "yy";
						break;
					case "long":
						f += "yyyy";
						break;
					default:
						f += "yy";
						break;
					}
				}
				break;
			case "am-pm":
				f += "AM/PM";
				break;
			case "week-of-year":
				;
			case "quarter":
				console.error("Excel does not support ODS format token " + i[3]);
				break;
			case "fill-character":
				if (i[1] === "/") {
					d = n.slice(u, Pt.lastIndex - i[0].length);
					f += '"' + d.replace(/"/g, '""') + '"*'
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					u = Pt.lastIndex
				}
				break;
			case "scientific-number":
				o = rt(i[0], false);
				f += "0." + yr("0", +o["min-decimal-places"] || +o["decimal-places"] || 2) + yr("?", +o["decimal-places"] - +o["min-decimal-places"] || 0) + "E" + (pt(o["forced-exponent-sign"]) ? "+": "") + yr("0", +o["min-exponent-digits"] || 2);
				break;
			case "fraction":
				o = rt(i[0], false);
				if (!+o["min-integer-digits"]) f += "#";
				else f += yr("0", +o["min-integer-digits"]);
				f += " ";
				f += yr("?", +o["min-numerator-digits"] || 1);
				f += "/";
				if ( + o["denominator-value"]) f += o["denominator-value"];
				else f += yr("?", +o["min-denominator-digits"] || 1);
				break;
			case "currency-symbol":
				if (i[1] === "/") {
					f += '"' + n.slice(u, Pt.lastIndex - i[0].length).replace(/"/g, '""') + '"'
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					u = Pt.lastIndex
				} else f += "$";
				break;
			case "text-properties":
				o = rt(i[0], false);
				switch ((o["color"] || "").toLowerCase().replace("#", "")) {
				case "ff0000":
					;
				case "red":
					f = "[Red]" + f;
					break;
				}
				break;
			case "text-content":
				f += "@";
				break;
			case "map":
				o = rt(i[0], false);
				if (it(o["condition"]) == "value()>=0") f = a[o["apply-style-name"]] + ";" + f;
				else console.error("ODS number format may be incorrect: " + o["condition"]);
				break;
			case "number":
				if (i[1] === "/") break;
				o = rt(i[0], false);
				l = "";
				l += yr("0", +o["min-integer-digits"] || 1);
				if (pt(o["grouping"])) l = he(yr("#", Math.max(0, 4 - l.length)) + l);
				if ( + o["min-decimal-places"] || +o["decimal-places"]) l += ".";
				if ( + o["min-decimal-places"]) l += yr("0", +o["min-decimal-places"] || 1);
				if ( + o["decimal-places"] - ( + o["min-decimal-places"] || 0)) l += yr("0", +o["decimal-places"] - ( + o["min-decimal-places"] || 0));
				f += l;
				break;
			case "embedded-text":
				if (i[1] === "/") {
					if (c == 0) f += '"' + n.slice(u, Pt.lastIndex - i[0].length).replace(/"/g, '""') + '"';
					else f = f.slice(0, c) + '"' + n.slice(u, Pt.lastIndex - i[0].length).replace(/"/g, '""') + '"' + f.slice(c)
				} else if (i[0].charAt(i[0].length - 2) !== "/") {
					u = Pt.lastIndex;
					c = -+rt(i[0], false)["position"] || 0
				}
				break;
			}
		}
		return a
	}
	function Ug(e, r, t) {
		var a = r || {};
		if (g != null && a.dense == null) a.dense = g;
		var n = Dt(e);
		var i = [],
		s;
		var f;
		var l, o = "",
		c = 0;
		var u;
		var h;
		var d = {},
		v = [];
		var p = {};
		if (a.dense) p["!data"] = [];
		var m, b;
		var w = {
			value: ""
		};
		var k = "",
		T = 0,
		y, E = "",
		_ = 0;
		var S = [],
		x = [];
		var A = -1,
		C = -1,
		R = {
			s: {
				r: 1e6,
				c: 1e7
			},
			e: {
				r: 0,
				c: 0
			}
		};
		var O = 0;
		var I = t || {},
		N = {};
		var F = [],
		D = {},
		P = 0,
		L = 0;
		var M = [],
		U = 1,
		B = 1;
		var W = [];
		var z = {
			Names: [],
			WBProps: {}
		};
		var H = {};
		var V = ["", ""];
		var $ = [],
		X = {};
		var G = "",
		j = 0;
		var K = false,
		Y = false;
		var Z = 0;
		Pt.lastIndex = 0;
		n = n.replace(/<!--([\s\S]*?)-->/gm, "").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm, "");
		while (m = Pt.exec(n)) switch (m[3] = m[3].replace(/_.*$/, "")) {
		case "table":
			;
		case "工作表":
			if (m[1] === "/") {
				if (R.e.c >= R.s.c && R.e.r >= R.s.r) p["!ref"] = Va(R);
				else p["!ref"] = "A1:A1";
				if (a.sheetRows > 0 && a.sheetRows <= R.e.r) {
					p["!fullref"] = p["!ref"];
					R.e.r = a.sheetRows - 1;
					p["!ref"] = Va(R)
				}
				if (F.length) p["!merges"] = F;
				if (M.length) p["!rows"] = M;
				u.name = u["名称"] || u.name;
				if (typeof JSON !== "undefined") JSON.stringify(u);
				v.push(u.name);
				d[u.name] = p;
				Y = false
			} else if (m[0].charAt(m[0].length - 2) !== "/") {
				u = rt(m[0], false);
				A = C = -1;
				R.s.r = R.s.c = 1e7;
				R.e.r = R.e.c = 0;
				p = {};
				if (a.dense) p["!data"] = [];
				F = [];
				M = [];
				Y = true
			}
			break;
		case "table-row-group":
			if (m[1] === "/")--O;
			else++O;
			break;
		case "table-row":
			;
		case "行":
			if (m[1] === "/") {
				A += U;
				U = 1;
				break
			}
			h = rt(m[0], false);
			if (h["行号"]) A = h["行号"] - 1;
			else if (A == -1) A = 0;
			U = +h["number-rows-repeated"] || 1;
			if (U < 10) for (Z = 0; Z < U; ++Z) if (O > 0) M[A + Z] = {
				level: O
			};
			C = -1;
			break;
		case "covered-table-cell":
			if (m[1] !== "/")++C;
			if (a.sheetStubs) {
				if (a.dense) {
					if (!p["!data"][A]) p["!data"][A] = [];
					p["!data"][A][C] = {
						t: "z"
					}
				} else p[za({
					r: A,
					c: C
				})] = {
					t: "z"
				}
			}
			k = "";
			S = [];
			break;
		case "table-cell":
			;
		case "数据":
			if (m[0].charAt(m[0].length - 2) === "/") {++C;
				w = rt(m[0], false);
				B = parseInt(w["number-columns-repeated"] || "1", 10);
				b = {
					t: "z",
					v: null
				};
				if (w.formula && a.cellFormula != false) b.f = hv(it(w.formula));
				if (w["style-name"] && N[w["style-name"]]) b.z = N[w["style-name"]];
				if ((w["数据类型"] || w["value-type"]) == "string") {
					b.t = "s";
					b.v = it(w["string-value"] || "");
					if (a.dense) {
						if (!p["!data"][A]) p["!data"][A] = [];
						p["!data"][A][C] = b
					} else {
						p[La(C) + Na(A)] = b
					}
				}
				C += B - 1
			} else if (m[1] !== "/") {++C;
				k = E = "";
				T = _ = 0;
				S = [];
				x = [];
				B = 1;
				var J = U ? A + U - 1 : A;
				if (C > R.e.c) R.e.c = C;
				if (C < R.s.c) R.s.c = C;
				if (A < R.s.r) R.s.r = A;
				if (J > R.e.r) R.e.r = J;
				w = rt(m[0], false);
				$ = [];
				X = {};
				b = {
					t: w["数据类型"] || w["value-type"],
					v: null
				};
				if (w["style-name"] && N[w["style-name"]]) b.z = N[w["style-name"]];
				if (a.cellFormula) {
					if (w.formula) w.formula = it(w.formula);
					if (w["number-matrix-columns-spanned"] && w["number-matrix-rows-spanned"]) {
						P = parseInt(w["number-matrix-rows-spanned"], 10) || 0;
						L = parseInt(w["number-matrix-columns-spanned"], 10) || 0;
						D = {
							s: {
								r: A,
								c: C
							},
							e: {
								r: A + P - 1,
								c: C + L - 1
							}
						};
						b.F = Va(D);
						W.push([D, b.F])
					}
					if (w.formula) b.f = hv(w.formula);
					else for (Z = 0; Z < W.length; ++Z) if (A >= W[Z][0].s.r && A <= W[Z][0].e.r) if (C >= W[Z][0].s.c && C <= W[Z][0].e.c) b.F = W[Z][1]
				}
				if (w["number-columns-spanned"] || w["number-rows-spanned"]) {
					P = parseInt(w["number-rows-spanned"], 10) || 0;
					L = parseInt(w["number-columns-spanned"], 10) || 0;
					D = {
						s: {
							r: A,
							c: C
						},
						e: {
							r: A + P - 1,
							c: C + L - 1
						}
					};
					F.push(D)
				}
				if (w["number-columns-repeated"]) B = parseInt(w["number-columns-repeated"], 10);
				switch (b.t) {
				case "boolean":
					b.t = "b";
					b.v = pt(w["boolean-value"]) || +w["boolean-value"] >= 1;
					break;
				case "float":
					b.t = "n";
					b.v = parseFloat(w.value);
					if (a.cellDates && b.z && Le(b.z)) {
						b.v = vr(b.v + (z.WBProps.date1904 ? 1462 : 0));
						b.t = typeof b.v == "number" ? "n": "d"
					}
					break;
				case "percentage":
					b.t = "n";
					b.v = parseFloat(w.value);
					break;
				case "currency":
					b.t = "n";
					b.v = parseFloat(w.value);
					break;
				case "date":
					b.t = "d";
					b.v = wr(w["date-value"], z.WBProps.date1904);
					if (!a.cellDates) {
						b.t = "n";
						b.v = dr(b.v, z.WBProps.date1904)
					}
					if (!b.z) b.z = "m/d/yy";
					break;
				case "time":
					b.t = "n";
					b.v = pr(w["time-value"]) / 86400;
					if (a.cellDates) {
						b.v = vr(b.v);
						b.t = typeof b.v == "number" ? "n": "d"
					}
					if (!b.z) b.z = "HH:MM:SS";
					break;
				case "number":
					b.t = "n";
					b.v = parseFloat(w["数据数值"]);
					break;
				default:
					if (b.t === "string" || b.t === "text" || !b.t) {
						b.t = "s";
						if (w["string-value"] != null) {
							k = it(w["string-value"]);
							S = []
						}
					} else throw new Error("Unsupported value type " + b.t);
				}
			} else {
				K = false;
				if (b.t === "s") {
					b.v = k || "";
					if (S.length) b.R = S;
					K = T == 0
				}
				if (H.Target) b.l = H;
				if ($.length > 0) {
					b.c = $;
					$ = []
				}
				if (k && a.cellText !== false) b.w = k;
				if (K) {
					b.t = "z";
					delete b.v
				}
				if (!K || a.sheetStubs) {
					if (! (a.sheetRows && a.sheetRows <= A)) {
						for (var q = 0; q < U; ++q) {
							B = parseInt(w["number-columns-repeated"] || "1", 10);
							if (a.dense) {
								if (!p["!data"][A + q]) p["!data"][A + q] = [];
								p["!data"][A + q][C] = q == 0 ? b: Tr(b);
								while (--B > 0) p["!data"][A + q][C + B] = Tr(b)
							} else {
								p[za({
									r: A + q,
									c: C
								})] = b;
								while (--B > 0) p[za({
									r: A + q,
									c: C + B
								})] = Tr(b)
							}
							if (R.e.c <= C) R.e.c = C
						}
					}
				}
				B = parseInt(w["number-columns-repeated"] || "1", 10);
				C += B - 1;
				B = 0;
				b = {};
				k = "";
				S = []
			}
			H = {};
			break;
		case "document":
			;
		case "document-content":
			;
		case "电子表格文档":
			;
		case "spreadsheet":
			;
		case "主体":
			;
		case "scripts":
			;
		case "styles":
			;
		case "font-face-decls":
			;
		case "master-styles":
			if (m[1] === "/") {
				if ((s = i.pop())[0] !== m[3]) throw "Bad state: " + s
			} else if (m[0].charAt(m[0].length - 2) !== "/") i.push([m[3], true]);
			break;
		case "annotation":
			if (m[1] === "/") {
				if ((s = i.pop())[0] !== m[3]) throw "Bad state: " + s;
				X.t = k;
				if (S.length) X.R = S;
				X.a = G;
				$.push(X);
				k = E;
				T = _;
				S = x
			} else if (m[0].charAt(m[0].length - 2) !== "/") {
				i.push([m[3], false]);
				var Q = rt(m[0], true);
				if (! (Q["display"] && pt(Q["display"]))) $.hidden = true;
				E = k;
				_ = T;
				x = S;
				k = "";
				T = 0;
				S = []
			}
			G = "";
			j = 0;
			break;
		case "creator":
			if (m[1] === "/") {
				G = n.slice(j, m.index)
			} else j = m.index + m[0].length;
			break;
		case "meta":
			;
		case "元数据":
			;
		case "settings":
			;
		case "config-item-set":
			;
		case "config-item-map-indexed":
			;
		case "config-item-map-entry":
			;
		case "config-item-map-named":
			;
		case "shapes":
			;
		case "frame":
			;
		case "text-box":
			;
		case "image":
			;
		case "data-pilot-tables":
			;
		case "list-style":
			;
		case "form":
			;
		case "dde-links":
			;
		case "event-listeners":
			;
		case "chart":
			if (m[1] === "/") {
				if ((s = i.pop())[0] !== m[3]) throw "Bad state: " + s
			} else if (m[0].charAt(m[0].length - 2) !== "/") i.push([m[3], false]);
			k = "";
			T = 0;
			S = [];
			break;
		case "scientific-number":
			;
		case "currency-symbol":
			;
		case "fill-character":
			break;
		case "text-style":
			;
		case "boolean-style":
			;
		case "number-style":
			;
		case "currency-style":
			;
		case "percentage-style":
			;
		case "date-style":
			;
		case "time-style":
			if (m[1] === "/") {
				var ee = Pt.lastIndex;
				Mg(n.slice(l, Pt.lastIndex), r, I);
				Pt.lastIndex = ee
			} else if (m[0].charAt(m[0].length - 2) !== "/") {
				l = Pt.lastIndex - m[0].length
			}
			break;
		case "script":
			break;
		case "libraries":
			break;
		case "automatic-styles":
			break;
		case "default-style":
			;
		case "page-layout":
			break;
		case "style":
			{
				var re = rt(m[0], false);
				if (re["family"] == "table-cell" && I[re["data-style-name"]]) N[re["name"]] = I[re["data-style-name"]]
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
			;
		case "month":
			;
		case "year":
			;
		case "era":
			;
		case "day-of-week":
			;
		case "week-of-year":
			;
		case "quarter":
			;
		case "hours":
			;
		case "minutes":
			;
		case "seconds":
			;
		case "am-pm":
			break;
		case "boolean":
			break;
		case "text":
			if (m[0].slice(-2) === "/>") break;
			else if (m[1] === "/") switch (i[i.length - 1][0]) {
			case "number-style":
				;
			case "date-style":
				;
			case "time-style":
				o += n.slice(c, m.index);
				break;
			} else c = m.index + m[0].length;
			break;
		case "named-range":
			f = rt(m[0], false);
			V = vv(f["cell-range-address"]);
			var te = {
				Name: f.name,
				Ref: V[0] + "!" + V[1]
			};
			if (Y) te.Sheet = v.length;
			z.Names.push(te);
			break;
		case "text-content":
			break;
		case "text-properties":
			break;
		case "embedded-text":
			break;
		case "body":
			;
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
			f = rt(m[0], false);
			switch (f["date-value"]) {
			case "1904-01-01":
				z.WBProps.date1904 = true;
				break;
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
			;
		case "文本串":
			if (["master-styles"].indexOf(i[i.length - 1][0]) > -1) break;
			if (m[1] === "/" && (!w || !w["string-value"])) {
				var ae = Lg(n.slice(T, m.index), y);
				k = (k.length > 0 ? k + "\n": "") + ae[0]
			} else if (m[0].slice(-2) == "/>") {
				k += "\n"
			} else {
				y = rt(m[0], false);
				T = m.index + m[0].length
			}
			break;
		case "s":
			break;
		case "database-range":
			if (m[1] === "/") break;
			try {
				V = vv(rt(m[0])["target-range-address"]);
				d[V[0]]["!autofilter"] = {
					ref: V[1]
				}
			} catch(ne) {}
			break;
		case "date":
			break;
		case "object":
			break;
		case "title":
			;
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
			;
		case "sender-lastname":
			;
		case "sender-initials":
			;
		case "sender-title":
			;
		case "sender-position":
			;
		case "sender-email":
			;
		case "sender-phone-private":
			;
		case "sender-fax":
			;
		case "sender-company":
			;
		case "sender-phone-work":
			;
		case "sender-street":
			;
		case "sender-city":
			;
		case "sender-postal-code":
			;
		case "sender-country":
			;
		case "sender-state-or-province":
			;
		case "author-name":
			;
		case "author-initials":
			;
		case "chapter":
			;
		case "file-name":
			;
		case "template-name":
			;
		case "sheet-name":
			break;
		case "event-listener":
			break;
		case "initial-creator":
			;
		case "creation-date":
			;
		case "print-date":
			;
		case "generator":
			;
		case "document-statistic":
			;
		case "user-defined":
			;
		case "editing-duration":
			;
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
			;
		case "source-cell-range":
			;
		case "source-service":
			;
		case "data-pilot-field":
			;
		case "data-pilot-level":
			;
		case "data-pilot-subtotals":
			;
		case "data-pilot-subtotal":
			;
		case "data-pilot-members":
			;
		case "data-pilot-member":
			;
		case "data-pilot-display-info":
			;
		case "data-pilot-sort-info":
			;
		case "data-pilot-layout-info":
			;
		case "data-pilot-field-reference":
			;
		case "data-pilot-groups":
			;
		case "data-pilot-group":
			;
		case "data-pilot-group-member":
			break;
		case "rect":
			break;
		case "dde-connection-decls":
			;
		case "dde-connection-decl":
			;
		case "dde-link":
			;
		case "dde-source":
			break;
		case "properties":
			break;
		case "property":
			break;
		case "a":
			if (m[1] !== "/") {
				H = rt(m[0], false);
				if (!H.href) break;
				H.Target = it(H.href);
				delete H.href;
				if (H.Target.charAt(0) == "#" && H.Target.indexOf(".") > -1) {
					V = vv(H.Target.slice(1));
					H.Target = "#" + V[0] + "!" + V[1]
				} else if (H.Target.match(/^\.\.[\\\/]/)) H.Target = H.Target.slice(3)
			}
			break;
		case "table-protection":
			break;
		case "data-pilot-grand-total":
			break;
		case "office-document-common-attrs":
			break;
		default:
			switch (m[2]) {
			case "dc:":
				;
			case "calcext:":
				;
			case "loext:":
				;
			case "ooo:":
				;
			case "chartooo:":
				;
			case "draw:":
				;
			case "style:":
				;
			case "chart:":
				;
			case "form:":
				;
			case "uof:":
				;
			case "表:":
				;
			case "字:":
				break;
			default:
				if (a.WTF) throw new Error(m);
			};
		}
		var ie = {
			Sheets: d,
			SheetNames: v,
			Workbook: z
		};
		if (a.bookSheets) delete ie.Sheets;
		return ie
	}
	function Bg(e, r) {
		r = r || {};
		if (Ur(e, "META-INF/manifest.xml")) mi(Wr(e, "META-INF/manifest.xml"), r);
		var t = zr(e, "styles.xml");
		var a = t && Mg(kt(t), r);
		var n = zr(e, "content.xml");
		if (!n) throw new Error("Missing content.xml in ODS / UOF file");
		var i = Ug(kt(n), r, a);
		if (Ur(e, "meta.xml")) i.Props = _i(Wr(e, "meta.xml"));
		i.bookType = "ods";
		return i
	}
	function Wg(e, r) {
		var t = Ug(e, r);
		t.bookType = "fods";
		return t
	}
	var zg = function() {
		var e = ["<office:master-styles>", '<style:master-page style:name="mp1" style:page-layout-name="mp1">', "<style:header/>", '<style:header-left style:display="false"/>', "<style:footer/>", '<style:footer-left style:display="false"/>', "</style:master-page>", "</office:master-styles>"].join("");
		var r = "<office:document-styles " + Ot({
			"xmlns:office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
			"xmlns:table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
			"xmlns:style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
			"xmlns:text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
			"xmlns:draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
			"xmlns:fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
			"xmlns:xlink": "http://www.w3.org/1999/xlink",
			"xmlns:dc": "http://purl.org/dc/elements/1.1/",
			"xmlns:number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
			"xmlns:svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
			"xmlns:of": "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
			"office:version": "1.2"
		}) + ">" + e + "</office:document-styles>";
		return function t() {
			return Kr + r
		}
	} ();
	function Hg(e, r) {
		var t = "number",
		a = "",
		n = {
			"style:name": r
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
					a += "<number:fill-character>" + lt(i.replace(/""/g, '"')) + "</number:fill-character>"
				} else {
					a += "<number:text>" + lt(i.replace(/""/g, '"')) + "</number:text>"
				}
				e = e.slice(s + 1);
				s = 0
			}
			var f = e.match(/# (\?+)\/(\?+)/);
			if (f) {
				a += It("number:fraction", null, {
					"number:min-integer-digits": 0,
					"number:min-numerator-digits": f[1].length,
					"number:max-denominator-value": Math.max( + f[1].replace(/./g, "9"), +f[2].replace(/./g, "9"))
				});
				break e
			}
			if (f = e.match(/# (\?+)\/(\d+)/)) {
				a += It("number:fraction", null, {
					"number:min-integer-digits": 0,
					"number:min-numerator-digits": f[1].length,
					"number:denominator-value": +f[2]
				});
				break e
			}
			if (f = e.match(/(\d+)(|\.\d+)%/)) {
				t = "percentage";
				a += It("number:number", null, {
					"number:decimal-places": f[2] && f.length - 1 || 0,
					"number:min-decimal-places": f[2] && f.length - 1 || 0,
					"number:min-integer-digits": f[1].length
				}) + "<number:text>%</number:text>";
				break e
			}
			var l = false;
			if (["y", "m", "d"].indexOf(e[0]) > -1) {
				t = "date";
				r: for (; s < e.length; ++s) switch (i = e[s].toLowerCase()) {
				case "h":
					;
				case "s":
					l = true; --s;
					break r;
				case "m":
					t:
					for (var o = s + 1; o < e.length; ++o) switch (e[o]) {
					case "y":
						;
					case "d":
						break t;
					case "h":
						;
					case "s":
						l = true; --s;
						break r;
					};
				case "y":
					;
				case "d":
					while ((e[++s] || "").toLowerCase() == i[0]) i += i[0]; --s;
					switch (i) {
					case "y":
						;
					case "yy":
						a += "<number:year/>";
						break;
					case "yyy":
						;
					case "yyyy":
						a += '<number:year number:style="long"/>';
						break;
					case "mmmmm":
						console.error("ODS has no equivalent of format |mmmmm|");
					case "m":
						;
					case "mm":
						;
					case "mmm":
						;
					case "mmmm":
						a += '<number:month number:style="' + (i.length % 2 ? "short": "long") + '" number:textual="' + (i.length >= 3 ? "true": "false") + '"/>';
						break;
					case "d":
						;
					case "dd":
						a += '<number:day number:style="' + (i.length % 2 ? "short": "long") + '"/>';
						break;
					case "ddd":
						;
					case "dddd":
						a += '<number:day-of-week number:style="' + (i.length % 2 ? "short": "long") + '"/>';
						break;
					}
					break;
				case '"':
					while (e[++s] != '"' || e[++s] == '"') i += e[s]; --s;
					a += "<number:text>" + lt(i.slice(1).replace(/""/g, '"')) + "</number:text>";
					break;
				case "\\":
					i = e[++s];
					a += "<number:text>" + lt(i) + "</number:text>";
					break;
				case "/":
					;
				case ":":
					a += "<number:text>" + lt(i) + "</number:text>";
					break;
				default:
					console.error("unrecognized character " + i + " in ODF format " + e);
				}
				if (!l) break e;
				e = e.slice(s + 1);
				s = 0
			}
			if (e.match(/^\[?[hms]/)) {
				if (t == "number") t = "time";
				if (e.match(/\[/)) {
					e = e.replace(/[\[\]]/g, "");
					n["number:truncate-on-overflow"] = "false"
				}
				for (; s < e.length; ++s) switch (i = e[s].toLowerCase()) {
				case "h":
					;
				case "m":
					;
				case "s":
					while ((e[++s] || "").toLowerCase() == i[0]) i += i[0]; --s;
					switch (i) {
					case "h":
						;
					case "hh":
						a += '<number:hours number:style="' + (i.length % 2 ? "short": "long") + '"/>';
						break;
					case "m":
						;
					case "mm":
						a += '<number:minutes number:style="' + (i.length % 2 ? "short": "long") + '"/>';
						break;
					case "s":
						;
					case "ss":
						if (e[s + 1] == ".") do {
							i += e[s + 1]; ++s
						} while ( e [ s + 1 ] == "0");
						a += '<number:seconds number:style="' + (i.match("ss") ? "long": "short") + '"' + (i.match(/\./) ? ' number:decimal-places="' + (i.match(/0+/) || [""])[0].length + '"': "") + "/>";
						break;
					}
					break;
				case '"':
					while (e[++s] != '"' || e[++s] == '"') i += e[s]; --s;
					a += "<number:text>" + lt(i.slice(1).replace(/""/g, '"')) + "</number:text>";
					break;
				case "/":
					;
				case ":":
					a += "<number:text>" + lt(i) + "</number:text>";
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
					};
				default:
					console.error("unrecognized character " + i + " in ODF format " + e);
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
					a += "<number:fill-character>" + lt(i.replace(/""/g, '"')) + "</number:fill-character>"
				} else {
					a += "<number:text>" + lt(i.replace(/""/g, '"')) + "</number:text>"
				}
				e = e.slice(s + 1);
				s = 0
			}
			var c = e.match(/([#0][0#,]*)(\.[0#]*|)(E[+]?0*|)/i);
			if (!c || !c[0]) console.error("Could not find numeric part of " + e);
			else {
				var u = c[1].replace(/,/g, "");
				a += "<number:" + (c[3] ? "scientific-": "") + "number" + ' number:min-integer-digits="' + (u.indexOf("0") == -1 ? "0": u.length - u.indexOf("0")) + '"' + (c[0].indexOf(",") > -1 ? ' number:grouping="true"': "") + (c[2] && ' number:decimal-places="' + (c[2].length - 1) + '"' || ' number:decimal-places="0"') + (c[3] && c[3].indexOf("+") > -1 ? ' number:forced-exponent-sign="true"': "") + (c[3] ? ' number:min-exponent-digits="' + c[3].match(/0+/)[0].length + '"': "") + ">" + "</number:" + (c[3] ? "scientific-": "") + "number>";
				s = c.index + c[0].length
			}
			if (e[s] == '"') {
				i = "";
				while (e[++s] != '"' || e[++s] == '"') i += e[s]; --s;
				a += "<number:text>" + lt(i.replace(/""/g, '"')) + "</number:text>"
			}
		}
		if (!a) {
			console.error("Could not generate ODS number format for |" + e + "|");
			return ""
		}
		return It("number:" + t + "-style", a, n)
	}
	function Vg(e, r, t) {
		var a = [];
		for (var n = 0; n < e.length; ++n) {
			var i = e[n];
			if (!i) continue;
			if (i.Sheet == (t == -1 ? null: t)) a.push(i)
		}
		if (!a.length) return "";
		return "      <table:named-expressions>\n" + a.map(function(e) {
			var r = (t == -1 ? "$": "") + pv(e.Ref);
			return "        " + It("table:named-range", null, {
				"table:name": e.Name,
				"table:cell-range-address": r,
				"table:base-cell-address": r.replace(/[\.]?[^\.]*$/, ".$A$1")
			})
		}).join("\n") + "\n      </table:named-expressions>\n"
	}
	var $g = function() {
		var e = function(e) {
			return lt(e).replace(/  +/g,
			function(e) {
				return '<text:s text:c="' + e.length + '"/>'
			}).replace(/\t/g, "<text:tab/>").replace(/\n/g, "</text:p><text:p>").replace(/^ /, "<text:s/>").replace(/ $/, "<text:s/>")
		};
		var r = "          <table:table-cell />\n";
		var t = function(t, a, n, i, s, f) {
			var l = [];
			l.push('      <table:table table:name="' + lt(a.SheetNames[n]) + '" table:style-name="ta1">\n');
			var o = 0,
			c = 0,
			u = Ha(t["!ref"] || "A1");
			var h = t["!merges"] || [],
			d = 0;
			var v = t["!data"] != null;
			if (t["!cols"]) {
				for (c = 0; c <= u.e.c; ++c) l.push("        <table:table-column" + (t["!cols"][c] ? ' table:style-name="co' + t["!cols"][c].ods + '"': "") + "></table:table-column>\n")
			}
			var p = "",
			m = t["!rows"] || [];
			for (o = 0; o < u.s.r; ++o) {
				p = m[o] ? ' table:style-name="ro' + m[o].ods + '"': "";
				l.push("        <table:table-row" + p + "></table:table-row>\n")
			}
			for (; o <= u.e.r; ++o) {
				p = m[o] ? ' table:style-name="ro' + m[o].ods + '"': "";
				l.push("        <table:table-row" + p + ">\n");
				for (c = 0; c < u.s.c; ++c) l.push(r);
				for (; c <= u.e.c; ++c) {
					var b = false,
					g = {},
					w = "";
					for (d = 0; d != h.length; ++d) {
						if (h[d].s.c > c) continue;
						if (h[d].s.r > o) continue;
						if (h[d].e.c < c) continue;
						if (h[d].e.r < o) continue;
						if (h[d].s.c != c || h[d].s.r != o) b = true;
						g["table:number-columns-spanned"] = h[d].e.c - h[d].s.c + 1;
						g["table:number-rows-spanned"] = h[d].e.r - h[d].s.r + 1;
						break
					}
					if (b) {
						l.push("          <table:covered-table-cell/>\n");
						continue
					}
					var k = za({
						r: o,
						c: c
					}),
					T = v ? (t["!data"][o] || [])[c] : t[k];
					if (T && T.f) {
						g["table:formula"] = lt(dv(T.f));
						if (T.F) {
							if (T.F.slice(0, k.length) == k) {
								var y = Ha(T.F);
								g["table:number-matrix-columns-spanned"] = y.e.c - y.s.c + 1;
								g["table:number-matrix-rows-spanned"] = y.e.r - y.s.r + 1
							}
						}
					}
					if (!T) {
						l.push(r);
						continue
					}
					switch (T.t) {
					case "b":
						w = T.v ? "TRUE": "FALSE";
						g["office:value-type"] = "boolean";
						g["office:boolean-value"] = T.v ? "true": "false";
						break;
					case "n":
						w = T.w || String(T.v || 0);
						g["office:value-type"] = "float";
						g["office:value"] = T.v || 0;
						break;
					case "s":
						;
					case "str":
						w = T.v == null ? "": T.v;
						g["office:value-type"] = "string";
						break;
					case "d":
						w = T.w || wr(T.v, f).toISOString();
						g["office:value-type"] = "date";
						g["office:date-value"] = wr(T.v, f).toISOString();
						g["table:style-name"] = "ce1";
						break;
					default:
						l.push(r);
						continue;
					}
					var E = e(w);
					if (T.l && T.l.Target) {
						var _ = T.l.Target;
						_ = _.charAt(0) == "#" ? "#" + pv(_.slice(1)) : _;
						if (_.charAt(0) != "#" && !_.match(/^\w+:/)) _ = "../" + _;
						E = It("text:a", E, {
							"xlink:href": _.replace(/&/g, "&amp;")
						})
					}
					if (s[T.z]) g["table:style-name"] = "ce" + s[T.z].slice(1);
					var S = It("text:p", E, {});
					if (T.c) {
						var x = "",
						A = "",
						C = {};
						for (var R = 0; R < T.c.length; ++R) {
							if (!x && T.c[R].a) x = T.c[R].a;
							A += "<text:p>" + e(T.c[R].t) + "</text:p>"
						}
						if (!T.c.hidden) C["office:display"] = true;
						S = It("office:annotation", A, C) + S
					}
					l.push("          " + It("table:table-cell", S, g) + "\n")
				}
				l.push("        </table:table-row>\n")
			}
			if ((a.Workbook || {}).Names) l.push(Vg(a.Workbook.Names, a.SheetNames, n));
			l.push("      </table:table>\n");
			return l.join("")
		};
		var a = function(e, r) {
			e.push(" <office:automatic-styles>\n");
			var t = 0;
			r.SheetNames.map(function(e) {
				return r.Sheets[e]
			}).forEach(function(r) {
				if (!r) return;
				if (r["!cols"]) {
					for (var a = 0; a < r["!cols"].length; ++a) if (r["!cols"][a]) {
						var n = r["!cols"][a];
						if (n.width == null && n.wpx == null && n.wch == null) continue;
						nc(n);
						n.ods = t;
						var i = r["!cols"][a].wpx + "px";
						e.push('  <style:style style:name="co' + t + '" style:family="table-column">\n');
						e.push('   <style:table-column-properties fo:break-before="auto" style:column-width="' + i + '"/>\n');
						e.push("  </style:style>\n"); ++t
					}
				}
			});
			var a = 0;
			r.SheetNames.map(function(e) {
				return r.Sheets[e]
			}).forEach(function(r) {
				if (!r) return;
				if (r["!rows"]) {
					for (var t = 0; t < r["!rows"].length; ++t) if (r["!rows"][t]) {
						r["!rows"][t].ods = a;
						var n = r["!rows"][t].hpx + "px";
						e.push('  <style:style style:name="ro' + a + '" style:family="table-row">\n');
						e.push('   <style:table-row-properties fo:break-before="auto" style:row-height="' + n + '"/>\n');
						e.push("  </style:style>\n"); ++a
					}
				}
			});
			e.push('  <style:style style:name="ta1" style:family="table" style:master-page-name="mp1">\n');
			e.push('   <style:table-properties table:display="true" style:writing-mode="lr-tb"/>\n');
			e.push("  </style:style>\n");
			e.push('  <number:date-style style:name="N37" number:automatic-order="true">\n');
			e.push('   <number:month number:style="long"/>\n');
			e.push("   <number:text>/</number:text>\n");
			e.push('   <number:day number:style="long"/>\n');
			e.push("   <number:text>/</number:text>\n");
			e.push("   <number:year/>\n");
			e.push("  </number:date-style>\n");
			var n = {};
			var i = 69;
			r.SheetNames.map(function(e) {
				return r.Sheets[e]
			}).forEach(function(r) {
				if (!r) return;
				var t = r["!data"] != null;
				if (!r["!ref"]) return;
				var a = Ha(r["!ref"]);
				for (var s = 0; s <= a.e.r; ++s) for (var f = 0; f <= a.e.c; ++f) {
					var l = t ? (r["!data"][s] || [])[f] : r[za({
						r: s,
						c: f
					})];
					if (!l || !l.z || l.z.toLowerCase() == "general") continue;
					if (!n[l.z]) {
						var o = Hg(l.z, "N" + i);
						if (o) {
							n[l.z] = "N" + i; ++i;
							e.push(o + "\n")
						}
					}
				}
			});
			e.push('  <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N37"/>\n');
			ir(n).forEach(function(r) {
				e.push('<style:style style:name="ce' + n[r].slice(1) + '" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="' + n[r] + '"/>\n')
			});
			e.push(" </office:automatic-styles>\n");
			return n
		};
		return function n(e, r) {
			var n = [Kr];
			var i = Ot({
				"xmlns:office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
				"xmlns:table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
				"xmlns:style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
				"xmlns:text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
				"xmlns:draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
				"xmlns:fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
				"xmlns:xlink": "http://www.w3.org/1999/xlink",
				"xmlns:dc": "http://purl.org/dc/elements/1.1/",
				"xmlns:meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
				"xmlns:number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
				"xmlns:presentation": "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
				"xmlns:svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
				"xmlns:chart": "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
				"xmlns:dr3d": "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
				"xmlns:math": "http://www.w3.org/1998/Math/MathML",
				"xmlns:form": "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
				"xmlns:script": "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
				"xmlns:ooo": "http://openoffice.org/2004/office",
				"xmlns:ooow": "http://openoffice.org/2004/writer",
				"xmlns:oooc": "http://openoffice.org/2004/calc",
				"xmlns:dom": "http://www.w3.org/2001/xml-events",
				"xmlns:xforms": "http://www.w3.org/2002/xforms",
				"xmlns:xsd": "http://www.w3.org/2001/XMLSchema",
				"xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
				"xmlns:sheet": "urn:oasis:names:tc:opendocument:sh33tjs:1.0",
				"xmlns:rpt": "http://openoffice.org/2005/report",
				"xmlns:of": "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
				"xmlns:xhtml": "http://www.w3.org/1999/xhtml",
				"xmlns:grddl": "http://www.w3.org/2003/g/data-view#",
				"xmlns:tableooo": "http://openoffice.org/2009/table",
				"xmlns:drawooo": "http://openoffice.org/2010/draw",
				"xmlns:calcext": "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
				"xmlns:loext": "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
				"xmlns:field": "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
				"xmlns:formx": "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0",
				"xmlns:css3t": "http://www.w3.org/TR/css3-text/",
				"office:version": "1.2"
			});
			var s = Ot({
				"xmlns:config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
				"office:mimetype": "application/vnd.oasis.opendocument.spreadsheet"
			});
			if (r.bookType == "fods") {
				n.push("<office:document" + i + s + ">\n");
				n.push(Ti().replace(/<office:document-meta.*?>/, "").replace(/<\/office:document-meta>/, "") + "\n")
			} else n.push("<office:document-content" + i + ">\n");
			var f = a(n, e);
			n.push("  <office:body>\n");
			n.push("    <office:spreadsheet>\n");
			if (((e.Workbook || {}).WBProps || {}).date1904) n.push('      <table:calculation-settings table:case-sensitive="false" table:search-criteria-must-apply-to-whole-cell="true" table:use-wildcards="true" table:use-regular-expressions="false" table:automatic-find-labels="false">\n        <table:null-date table:date-value="1904-01-01"/>\n      </table:calculation-settings>\n');
			for (var l = 0; l != e.SheetNames.length; ++l) n.push(t(e.Sheets[e.SheetNames[l]], e, l, r, f, ((e.Workbook || {}).WBProps || {}).date1904));
			if ((e.Workbook || {}).Names) n.push(Vg(e.Workbook.Names, e.SheetNames, -1));
			n.push("    </office:spreadsheet>\n");
			n.push("  </office:body>\n");
			if (r.bookType == "fods") n.push("</office:document>");
			else n.push("</office:document-content>");
			return n.join("")
		}
	} ();
	function Xg(e, r) {
		if (r.bookType == "fods") return $g(e, r);
		var t = Xr();
		var a = "";
		var n = [];
		var i = [];
		a = "mimetype";
		$r(t, a, "application/vnd.oasis.opendocument.spreadsheet");
		a = "content.xml";
		$r(t, a, $g(e, r));
		n.push([a, "text/xml"]);
		i.push([a, "ContentFile"]);
		a = "styles.xml";
		$r(t, a, zg(e, r));
		n.push([a, "text/xml"]);
		i.push([a, "StylesFile"]);
		a = "meta.xml";
		$r(t, a, Kr + Ti());
		n.push([a, "text/xml"]);
		i.push([a, "MetadataFile"]);
		a = "manifest.rdf";
		$r(t, a, ki(i));
		n.push([a, "application/rdf+xml"]);
		a = "META-INF/manifest.xml";
		$r(t, a, bi(n));
		return t
	}
	var Gg = function() {
		try {
			if (typeof Uint8Array == "undefined") return "slice";
			if (typeof Uint8Array.prototype.subarray == "undefined") return "slice";
			if (typeof Buffer !== "undefined") {
				if (typeof Buffer.prototype.subarray == "undefined") return "slice";
				if ((typeof Buffer.from == "function" ? Buffer.from([72, 62]) : new Buffer([72, 62])) instanceof Uint8Array) return "subarray";
				return "slice"
			}
			return "subarray"
		} catch(e) {
			return "slice"
		}
	} ();
	function jg(e) {
		return new DataView(e.buffer, e.byteOffset, e.byteLength)
	}
	function Kg(e) {
		return typeof TextDecoder != "undefined" ? (new TextDecoder).decode(e) : kt(N(e))
	}
	function Yg(e) {
		return typeof TextEncoder != "undefined" ? (new TextEncoder).encode(e) : O(Tt(e))
	}
	function Zg(e) {
		var r = 0;
		for (var t = 0; t < e.length; ++t) r += e[t].length;
		var a = new Uint8Array(r);
		var n = 0;
		for (t = 0; t < e.length; ++t) {
			var i = e[t],
			s = i.length;
			if (s < 250) {
				for (var f = 0; f < s; ++f) a[n++] = i[f]
			} else {
				a.set(i, n);
				n += s
			}
		}
		return a
	}
	function Jg(e) {
		e -= e >> 1 & 1431655765;
		e = (e & 858993459) + (e >> 2 & 858993459);
		return (e + (e >> 4) & 252645135) * 16843009 >>> 24
	}
	function qg(e, r) {
		var t = (e[r + 15] & 127) << 7 | e[r + 14] >> 1;
		var a = e[r + 14] & 1;
		for (var n = r + 13; n >= r; --n) a = a * 256 + e[n];
		return (e[r + 15] & 128 ? -a: a) * Math.pow(10, t - 6176)
	}
	function Qg(e, r, t) {
		var a = Math.floor(t == 0 ? 0 : Math.LOG10E * Math.log(Math.abs(t))) + 6176 - 16;
		var n = t / Math.pow(10, a - 6176);
		e[r + 15] |= a >> 7;
		e[r + 14] |= (a & 127) << 1;
		for (var i = 0; n >= 1; ++i, n /= 256) e[r + i] = n & 255;
		e[r + 15] |= t >= 0 ? 0 : 128
	}
	function ew(e, r) {
		var t = r.l;
		var a = e[t] & 127;
		e: if (e[t++] >= 128) {
			a |= (e[t] & 127) << 7;
			if (e[t++] < 128) break e;
			a |= (e[t] & 127) << 14;
			if (e[t++] < 128) break e;
			a |= (e[t] & 127) << 21;
			if (e[t++] < 128) break e;
			a += (e[t] & 127) * Math.pow(2, 28); ++t;
			if (e[t++] < 128) break e;
			a += (e[t] & 127) * Math.pow(2, 35); ++t;
			if (e[t++] < 128) break e;
			a += (e[t] & 127) * Math.pow(2, 42); ++t;
			if (e[t++] < 128) break e
		}
		r.l = t;
		return a
	}
	function rw(e) {
		var r = new Uint8Array(7);
		r[0] = e & 127;
		var t = 1;
		e: if (e > 127) {
			r[t - 1] |= 128;
			r[t] = e >> 7 & 127; ++t;
			if (e <= 16383) break e;
			r[t - 1] |= 128;
			r[t] = e >> 14 & 127; ++t;
			if (e <= 2097151) break e;
			r[t - 1] |= 128;
			r[t] = e >> 21 & 127; ++t;
			if (e <= 268435455) break e;
			r[t - 1] |= 128;
			r[t] = e / 256 >>> 21 & 127; ++t;
			if (e <= 34359738367) break e;
			r[t - 1] |= 128;
			r[t] = e / 65536 >>> 21 & 127; ++t;
			if (e <= 4398046511103) break e;
			r[t - 1] |= 128;
			r[t] = e / 16777216 >>> 21 & 127; ++t
		}
		return r[Gg](0, t)
	}
	function tw(e) {
		var r = {
			l: 0
		};
		var t = [];
		while (r.l < e.length) t.push(ew(e, r));
		return t
	}
	function aw(e) {
		return Zg(e.map(function(e) {
			return rw(e)
		}))
	}
	function nw(e) {
		var r = 0,
		t = e[r] & 127;
		if (e[r++] < 128) return t;
		t |= (e[r] & 127) << 7;
		if (e[r++] < 128) return t;
		t |= (e[r] & 127) << 14;
		if (e[r++] < 128) return t;
		t |= (e[r] & 127) << 21;
		if (e[r++] < 128) return t;
		t |= (e[r] & 15) << 28;
		return t
	}
	function iw(e) {
		var r = 0,
		t = e[r] & 127,
		a = 0;
		e: if (e[r++] >= 128) {
			t |= (e[r] & 127) << 7;
			if (e[r++] < 128) break e;
			t |= (e[r] & 127) << 14;
			if (e[r++] < 128) break e;
			t |= (e[r] & 127) << 21;
			if (e[r++] < 128) break e;
			t |= (e[r] & 127) << 28;
			a = e[r] >> 4 & 7;
			if (e[r++] < 128) break e;
			a |= (e[r] & 127) << 3;
			if (e[r++] < 128) break e;
			a |= (e[r] & 127) << 10;
			if (e[r++] < 128) break e;
			a |= (e[r] & 127) << 17;
			if (e[r++] < 128) break e;
			a |= (e[r] & 127) << 24;
			if (e[r++] < 128) break e;
			a |= (e[r] & 127) << 31
		}
		return [t >>> 0, a >>> 0]
	}
	function sw(e) {
		var r = [],
		t = {
			l: 0
		};
		while (t.l < e.length) {
			var a = t.l;
			var n = ew(e, t);
			var i = n & 7;
			n = n / 8 | 0;
			var s;
			var f = t.l;
			switch (i) {
			case 0:
				{
					while (e[f++] >= 128);
					s = e[Gg](t.l, f);
					t.l = f
				}
				break;
			case 1:
				{
					s = e[Gg](f, f + 8);
					t.l = f + 8
				}
				break;
			case 2:
				{
					var l = ew(e, t);
					s = e[Gg](t.l, t.l + l);
					t.l += l
				}
				break;
			case 5:
				{
					s = e[Gg](f, f + 4);
					t.l = f + 4
				}
				break;
			default:
				throw new Error("PB Type ".concat(i, " for Field ").concat(n, " at offset ").concat(a));
			}
			var o = {
				data: s,
				type: i
			};
			if (r[n] == null) r[n] = [];
			r[n].push(o)
		}
		return r
	}
	function fw(e) {
		var r = [];
		e.forEach(function(e, t) {
			if (t == 0) return;
			e.forEach(function(e) {
				if (!e.data) return;
				r.push(rw(t * 8 + e.type));
				if (e.type == 2) r.push(rw(e.data.length));
				r.push(e.data)
			})
		});
		return Zg(r)
	}
	function lw(e, r) {
		return (e == null ? void 0 : e.map(function(e) {
			return r(e.data)
		})) || []
	}
	function ow(e) {
		var r;
		var t = [],
		a = {
			l: 0
		};
		while (a.l < e.length) {
			var n = ew(e, a);
			var i = sw(e[Gg](a.l, a.l + n));
			a.l += n;
			var s = {
				id: nw(i[1][0].data),
				messages: []
			};
			i[2].forEach(function(r) {
				var t = sw(r.data);
				var n = nw(t[3][0].data);
				s.messages.push({
					meta: t,
					data: e[Gg](a.l, a.l + n)
				});
				a.l += n
			});
			if ((r = i[3]) == null ? void 0 : r[0]) s.merge = nw(i[3][0].data) >>> 0 > 0;
			t.push(s)
		}
		return t
	}
	function cw(e) {
		var r = [];
		e.forEach(function(e) {
			var t = [[], [{
				data: rw(e.id),
				type: 0
			}], []];
			if (e.merge != null) t[3] = [{
				data: rw( + !!e.merge),
				type: 0
			}];
			var a = [];
			e.messages.forEach(function(e) {
				a.push(e.data);
				e.meta[3] = [{
					type: 0,
					data: rw(e.data.length)
				}];
				t[2].push({
					data: fw(e.meta),
					type: 2
				})
			});
			var n = fw(t);
			r.push(rw(n.length));
			r.push(n);
			a.forEach(function(e) {
				return r.push(e)
			})
		});
		return Zg(r)
	}
	function uw(e, r) {
		if (e != 0) throw new Error("Unexpected Snappy chunk type ".concat(e));
		var t = {
			l: 0
		};
		var a = ew(r, t);
		var n = [];
		var i = t.l;
		while (i < r.length) {
			var s = r[i] & 3;
			if (s == 0) {
				var f = r[i++] >> 2;
				if (f < 60)++f;
				else {
					var l = f - 59;
					f = r[i];
					if (l > 1) f |= r[i + 1] << 8;
					if (l > 2) f |= r[i + 2] << 16;
					if (l > 3) f |= r[i + 3] << 24;
					f >>>= 0;
					f++;
					i += l
				}
				n.push(r[Gg](i, i + f));
				i += f;
				continue
			} else {
				var o = 0,
				c = 0;
				if (s == 1) {
					c = (r[i] >> 2 & 7) + 4;
					o = (r[i++] & 224) << 3;
					o |= r[i++]
				} else {
					c = (r[i++] >> 2) + 1;
					if (s == 2) {
						o = r[i] | r[i + 1] << 8;
						i += 2
					} else {
						o = (r[i] | r[i + 1] << 8 | r[i + 2] << 16 | r[i + 3] << 24) >>> 0;
						i += 4
					}
				}
				if (o == 0) throw new Error("Invalid offset 0");
				var u = n.length - 1,
				h = o;
				while (u >= 0 && h >= n[u].length) {
					h -= n[u].length; --u
				}
				if (u < 0) {
					if (h == 0) h = n[u = 0].length;
					else throw new Error("Invalid offset beyond length")
				}
				if (c < h) n.push(n[u][Gg](n[u].length - h, n[u].length - h + c));
				else {
					if (h > 0) {
						n.push(n[u][Gg](n[u].length - h));
						c -= h
					}++u;
					while (c >= n[u].length) {
						n.push(n[u]);
						c -= n[u].length; ++u
					}
					if (c) n.push(n[u][Gg](0, c))
				}
				if (n.length > 25) n = [Zg(n)]
			}
		}
		var d = 0;
		for (var v = 0; v < n.length; ++v) d += n[v].length;
		if (d != a) throw new Error("Unexpected length: ".concat(d, " != ").concat(a));
		return n
	}
	function hw(e) {
		if (Array.isArray(e)) e = new Uint8Array(e);
		var r = [];
		var t = 0;
		while (t < e.length) {
			var a = e[t++];
			var n = e[t] | e[t + 1] << 8 | e[t + 2] << 16;
			t += 3;
			r.push.apply(r, uw(a, e[Gg](t, t + n)));
			t += n
		}
		if (t !== e.length) throw new Error("data is not a valid framed stream!");
		return r.length == 1 ? r[0] : Zg(r)
	}
	function dw(e) {
		var r = [];
		var t = 0;
		while (t < e.length) {
			var a = Math.min(e.length - t, 268435455);
			var n = new Uint8Array(4);
			r.push(n);
			var i = rw(a);
			var s = i.length;
			r.push(i);
			if (a <= 60) {
				s++;
				r.push(new Uint8Array([a - 1 << 2]))
			} else if (a <= 256) {
				s += 2;
				r.push(new Uint8Array([240, a - 1 & 255]))
			} else if (a <= 65536) {
				s += 3;
				r.push(new Uint8Array([244, a - 1 & 255, a - 1 >> 8 & 255]))
			} else if (a <= 16777216) {
				s += 4;
				r.push(new Uint8Array([248, a - 1 & 255, a - 1 >> 8 & 255, a - 1 >> 16 & 255]))
			} else if (a <= 4294967296) {
				s += 5;
				r.push(new Uint8Array([252, a - 1 & 255, a - 1 >> 8 & 255, a - 1 >> 16 & 255, a - 1 >>> 24 & 255]))
			}
			r.push(e[Gg](t, t + a));
			s += a;
			n[0] = 0;
			n[1] = s & 255;
			n[2] = s >> 8 & 255;
			n[3] = s >> 16 & 255;
			t += a
		}
		return Zg(r)
	}
	var vw = function() {
		return {
			sst: [],
			rsst: [],
			ofmt: [],
			nfmt: [],
			fmla: [],
			ferr: [],
			cmnt: []
		}
	};
	function pw(e, r, t, a, n) {
		var i, s, f, l;
		var o = r & 255,
		c = r >> 8;
		var u = c >= 5 ? n: a;
		e: if (t & (c > 4 ? 8 : 4) && e.t == "n" && o == 7) {
			var h = ((i = u[7]) == null ? void 0 : i[0]) ? nw(u[7][0].data) : -1;
			if (h == -1) break e;
			var d = ((s = u[15]) == null ? void 0 : s[0]) ? nw(u[15][0].data) : -1;
			var v = ((f = u[16]) == null ? void 0 : f[0]) ? nw(u[16][0].data) : -1;
			var p = ((l = u[40]) == null ? void 0 : l[0]) ? nw(u[40][0].data) : -1;
			var m = e.v,
			b = m;
			r: if (p) {
				if (m == 0) {
					d = v = 2;
					break r
				}
				if (m >= 604800) d = 1;
				else if (m >= 86400) d = 2;
				else if (m >= 3600) d = 4;
				else if (m >= 60) d = 8;
				else if (m >= 1) d = 16;
				else d = 32;
				if (Math.floor(m) != m) v = 32;
				else if (m % 60) v = 16;
				else if (m % 3600) v = 8;
				else if (m % 86400) v = 4;
				else if (m % 604800) v = 2;
				if (v < d) v = d
			}
			if (d == -1 || v == -1) break e;
			var g = [],
			w = [];
			if (d == 1) {
				b = m / 604800;
				if (v == 1) {
					w.push('d"d"')
				} else {
					b |= 0;
					m -= 604800 * b
				}
				g.push(b + (h == 2 ? " week" + (b == 1 ? "": "s") : h == 1 ? "w": ""))
			}
			if (d <= 2 && v >= 2) {
				b = m / 86400;
				if (v > 2) {
					b |= 0;
					m -= 86400 * b
				}
				w.push('d"d"');
				g.push(b + (h == 2 ? " day" + (b == 1 ? "": "s") : h == 1 ? "d": ""))
			}
			if (d <= 4 && v >= 4) {
				b = m / 3600;
				if (v > 4) {
					b |= 0;
					m -= 3600 * b
				}
				w.push((d >= 4 ? "[h]": "h") + '"h"');
				g.push(b + (h == 2 ? " hour" + (b == 1 ? "": "s") : h == 1 ? "h": ""))
			}
			if (d <= 8 && v >= 8) {
				b = m / 60;
				if (v > 8) {
					b |= 0;
					m -= 60 * b
				}
				w.push((d >= 8 ? "[m]": "m") + '"m"');
				if (h == 0) g.push((d == 8 && v == 8 || b >= 10 ? "": "0") + b);
				else g.push(b + (h == 2 ? " minute" + (b == 1 ? "": "s") : h == 1 ? "m": ""))
			}
			if (d <= 16 && v >= 16) {
				b = m;
				if (v > 16) {
					b |= 0;
					m -= b
				}
				w.push((d >= 16 ? "[s]": "s") + '"s"');
				if (h == 0) g.push((v == 16 && d == 16 || b >= 10 ? "": "0") + b);
				else g.push(b + (h == 2 ? " second" + (b == 1 ? "": "s") : h == 1 ? "s": ""))
			}
			if (v >= 32) {
				b = Math.round(1e3 * m);
				if (d < 32) w.push('.000"ms"');
				if (h == 0) g.push((b >= 100 ? "": b >= 10 ? "0": "00") + b);
				else g.push(b + (h == 2 ? " millisecond" + (b == 1 ? "": "s") : h == 1 ? "ms": ""))
			}
			e.w = g.join(h == 0 ? ":": " ");
			e.z = w.join(h == 0 ? '":"': " ");
			if (h == 0) e.w = e.w.replace(/:(\d\d\d)$/, ".$1")
		}
	}
	function mw(e, r, t, a) {
		var n = jg(e);
		var i = n.getUint32(4, true);
		var s = -1,
		f = -1,
		l = -1,
		o = NaN,
		c = 0,
		u = new Date(Date.UTC(2001, 0, 1));
		var h = t > 1 ? 12 : 8;
		if (i & 2) {
			l = n.getUint32(h, true);
			h += 4
		}
		h += Jg(i & (t > 1 ? 3468 : 396)) * 4;
		if (i & 512) {
			s = n.getUint32(h, true);
			h += 4
		}
		h += Jg(i & (t > 1 ? 12288 : 4096)) * 4;
		if (i & 16) {
			f = n.getUint32(h, true);
			h += 4
		}
		if (i & 32) {
			o = n.getFloat64(h, true);
			h += 8
		}
		if (i & 64) {
			u.setTime(u.getTime() + (c = n.getFloat64(h, true)) * 1e3);
			h += 8
		}
		if (t > 1) {
			i = n.getUint32(8, true) >>> 16;
			if (i & 255) {
				if (l == -1) l = n.getUint32(h, true);
				h += 4
			}
		}
		var d;
		var v = e[t >= 4 ? 1 : 2];
		switch (v) {
		case 0:
			return void 0;
		case 2:
			d = {
				t: "n",
				v: o
			};
			break;
		case 3:
			d = {
				t: "s",
				v: r.sst[f]
			};
			break;
		case 5:
			{
				if (a == null ? void 0 : a.cellDates) d = {
					t: "d",
					v: u
				};
				else d = {
					t: "n",
					v: c / 86400 + 35430,
					z: q[14]
				}
			}
			break;
		case 6:
			d = {
				t: "b",
				v: o > 0
			};
			break;
		case 7:
			d = {
				t: "n",
				v: o
			};
			break;
		case 8:
			d = {
				t: "e",
				v: 0
			};
			break;
		case 9:
			{
				if (s > -1) {
					var p = r.rsst[s];
					d = {
						t: "s",
						v: p.v
					};
					if (p.l) d.l = {
						Target: p.l
					}
				} else throw new Error("Unsupported cell type ".concat(e[Gg](0, 4)))
			}
			break;
		default:
			throw new Error("Unsupported cell type ".concat(e[Gg](0, 4)));
		}
		if (l > -1) pw(d, v | t << 8, i, r.ofmt[l], r.nfmt[l]);
		if (v == 7) d.v /= 86400;
		return d
	}
	function bw(e, r, t) {
		var a = jg(e);
		var n = a.getUint32(4, true);
		var i = a.getUint32(8, true);
		var s = 12;
		var f = -1,
		l = -1,
		o = -1,
		c = NaN,
		u = NaN,
		h = 0,
		d = new Date(Date.UTC(2001, 0, 1)),
		v = -1,
		p = -1;
		if (i & 1) {
			c = qg(e, s);
			s += 16
		}
		if (i & 2) {
			u = a.getFloat64(s, true);
			s += 8
		}
		if (i & 4) {
			d.setTime(d.getTime() + (h = a.getFloat64(s, true)) * 1e3);
			s += 8
		}
		if (i & 8) {
			l = a.getUint32(s, true);
			s += 4
		}
		if (i & 16) {
			f = a.getUint32(s, true);
			s += 4
		}
		s += Jg(i & 480) * 4;
		if (i & 512) {
			p = a.getUint32(s, true);
			s += 4
		}
		s += Jg(i & 1024) * 4;
		if (i & 2048) {
			v = a.getUint32(s, true);
			s += 4
		}
		var m;
		var b = e[1];
		switch (b) {
		case 0:
			m = {
				t: "z"
			};
			break;
		case 2:
			m = {
				t: "n",
				v: c
			};
			break;
		case 3:
			m = {
				t: "s",
				v: r.sst[l]
			};
			break;
		case 5:
			{
				if (t == null ? void 0 : t.cellDates) m = {
					t: "d",
					v: d
				};
				else m = {
					t: "n",
					v: h / 86400 + 35430,
					z: q[14]
				}
			}
			break;
		case 6:
			m = {
				t: "b",
				v: u > 0
			};
			break;
		case 7:
			m = {
				t: "n",
				v: u
			};
			break;
		case 8:
			m = {
				t: "e",
				v: 0
			};
			break;
		case 9:
			{
				if (f > -1) {
					var g = r.rsst[f];
					m = {
						t: "s",
						v: g.v
					};
					if (g.l) m.l = {
						Target: g.l
					}
				} else throw new Error("Unsupported cell type ".concat(e[1], " : ").concat(i & 31, " : ").concat(e[Gg](0, 4)))
			}
			break;
		case 10:
			m = {
				t: "n",
				v: c
			};
			break;
		default:
			throw new Error("Unsupported cell type ".concat(e[1], " : ").concat(i & 31, " : ").concat(e[Gg](0, 4)));
		}
		s += Jg(i & 4096) * 4;
		if (i & 516096) {
			if (o == -1) o = a.getUint32(s, true);
			s += 4
		}
		if (i & 524288) {
			var w = a.getUint32(s, true);
			s += 4;
			if (r.cmnt[w]) m.c = Rw(r.cmnt[w])
		}
		if (o > -1) pw(m, b | 5 << 8, i >> 13, r.ofmt[o], r.nfmt[o]);
		if (b == 7) m.v /= 86400;
		return m
	}
	function gw(e, r) {
		var t = new Uint8Array(32),
		a = jg(t),
		n = 12,
		i = 0;
		t[0] = 5;
		switch (e.t) {
		case "n":
			if (e.z && Le(e.z)) {
				t[1] = 5;
				a.setFloat64(n, (vr(e.v + 1462).getTime() - Date.UTC(2001, 0, 1)) / 1e3, true);
				i |= 4;
				n += 8;
				break
			} else {
				t[1] = 2;
				Qg(t, n, e.v);
				i |= 1;
				n += 16
			}
			break;
		case "b":
			t[1] = 6;
			a.setFloat64(n, e.v ? 1 : 0, true);
			i |= 2;
			n += 8;
			break;
		case "s":
			{
				var s = e.v == null ? "": String(e.v);
				if (e.l) {
					var f = r.rsst.findIndex(function(r) {
						var t;
						return r.v == s && r.l == ((t = e.l) == null ? void 0 : t.Target)
					});
					if (f == -1) r.rsst[f = r.rsst.length] = {
						v: s,
						l: e.l.Target
					};
					t[1] = 9;
					a.setUint32(n, f, true);
					i |= 16;
					n += 4
				} else {
					var l = r.sst.indexOf(s);
					if (l == -1) r.sst[l = r.sst.length] = s;
					t[1] = 3;
					a.setUint32(n, l, true);
					i |= 8;
					n += 4
				}
			}
			break;
		case "d":
			t[1] = 5;
			a.setFloat64(n, (e.v.getTime() - Date.UTC(2001, 0, 1)) / 1e3, true);
			i |= 4;
			n += 8;
			break;
		case "z":
			t[1] = 0;
			break;
		default:
			throw "unsupported cell type " + e.t;
		}
		if (e.c) {
			r.cmnt.push(Ow(e.c));
			a.setUint32(n, r.cmnt.length - 1, true);
			i |= 524288;
			n += 4
		}
		a.setUint32(8, i, true);
		return t[Gg](0, n)
	}
	function ww(e, r) {
		var t = new Uint8Array(32),
		a = jg(t),
		n = 12,
		i = 0,
		s = "";
		t[0] = 4;
		switch (e.t) {
		case "n":
			break;
		case "b":
			break;
		case "s":
			{
				s = e.v == null ? "": String(e.v);
				if (e.l) {
					var f = r.rsst.findIndex(function(r) {
						var t;
						return r.v == s && r.l == ((t = e.l) == null ? void 0 : t.Target)
					});
					if (f == -1) r.rsst[f = r.rsst.length] = {
						v: s,
						l: e.l.Target
					};
					t[1] = 9;
					a.setUint32(n, f, true);
					i |= 512;
					n += 4
				} else {}
			}
			break;
		case "d":
			break;
		case "e":
			break;
		case "z":
			break;
		default:
			throw "unsupported cell type " + e.t;
		}
		if (e.c) {
			a.setUint32(n, r.cmnt.length - 1, true);
			i |= 4096;
			n += 4
		}
		switch (e.t) {
		case "n":
			t[1] = 2;
			a.setFloat64(n, e.v, true);
			i |= 32;
			n += 8;
			break;
		case "b":
			t[1] = 6;
			a.setFloat64(n, e.v ? 1 : 0, true);
			i |= 32;
			n += 8;
			break;
		case "s":
			{
				s = e.v == null ? "": String(e.v);
				if (e.l) {} else {
					var l = r.sst.indexOf(s);
					if (l == -1) r.sst[l = r.sst.length] = s;
					t[1] = 3;
					a.setUint32(n, l, true);
					i |= 16;
					n += 4
				}
			}
			break;
		case "d":
			t[1] = 5;
			a.setFloat64(n, (e.v.getTime() - Date.UTC(2001, 0, 1)) / 1e3, true);
			i |= 64;
			n += 8;
			break;
		case "z":
			t[1] = 0;
			break;
		default:
			throw "unsupported cell type " + e.t;
		}
		a.setUint32(8, i, true);
		return t[Gg](0, n)
	}
	function kw(e, r, t) {
		switch (e[0]) {
		case 0:
			;
		case 1:
			;
		case 2:
			;
		case 3:
			;
		case 4:
			return mw(e, r, e[0], t);
		case 5:
			return bw(e, r, t);
		default:
			throw new Error("Unsupported payload version ".concat(e[0]));
		}
	}
	function Tw(e) {
		var r = sw(e);
		return nw(r[1][0].data)
	}
	function yw(e) {
		return fw([[], [{
			type: 0,
			data: rw(e)
		}]])
	}
	function Ew(e, r) {
		var t;
		var a = ((t = e.messages[0].meta[5]) == null ? void 0 : t[0]) ? tw(e.messages[0].meta[5][0].data) : [];
		var n = a.indexOf(r);
		if (n == -1) {
			a.push(r);
			e.messages[0].meta[5] = [{
				type: 2,
				data: aw(a)
			}]
		}
	}
	function _w(e, r) {
		var t;
		var a = ((t = e.messages[0].meta[5]) == null ? void 0 : t[0]) ? tw(e.messages[0].meta[5][0].data) : [];
		e.messages[0].meta[5] = [{
			type: 2,
			data: aw(a.filter(function(e) {
				return e != r
			}))
		}]
	}
	function Sw(e, r) {
		var t = sw(r.data);
		var a = nw(t[1][0].data);
		var n = t[3];
		var i = []; (n || []).forEach(function(r) {
			var t, n;
			var s = sw(r.data);
			if (!s[1]) return;
			var f = nw(s[1][0].data) >>> 0;
			switch (a) {
			case 1:
				i[f] = Kg(s[3][0].data);
				break;
			case 8:
				{
					var l = e[Tw(s[9][0].data)][0];
					var o = sw(l.data);
					var c = e[Tw(o[1][0].data)][0];
					var u = nw(c.meta[1][0].data);
					if (u != 2001) throw new Error("2000 unexpected reference to ".concat(u));
					var h = sw(c.data);
					var d = {
						v: h[3].map(function(e) {
							return Kg(e.data)
						}).join("")
					};
					i[f] = d;
					e: if ((t = h == null ? void 0 : h[11]) == null ? void 0 : t[0]) {
						var v = (n = sw(h[11][0].data)) == null ? void 0 : n[1];
						if (!v) break e;
						v.forEach(function(r) {
							var t, a, n;
							var i = sw(r.data);
							if ((t = i[2]) == null ? void 0 : t[0]) {
								var s = e[Tw((a = i[2]) == null ? void 0 : a[0].data)][0];
								var f = nw(s.meta[1][0].data);
								switch (f) {
								case 2032:
									var l = sw(s.data);
									if (((n = l == null ? void 0 : l[2]) == null ? void 0 : n[0]) && !d.l) d.l = Kg(l[2][0].data);
									break;
								case 2039:
									break;
								default:
									console.log("unrecognized ObjectAttribute type ".concat(f));
								}
							}
						})
					}
				}
				break;
			case 2:
				i[f] = sw(s[6][0].data);
				break;
			case 3:
				i[f] = sw(s[5][0].data);
				break;
			case 10:
				{
					var p = e[Tw(s[10][0].data)][0];
					i[f] = Cw(e, p.data)
				}
				break;
			default:
				throw a;
			}
		});
		return i
	}
	function xw(e, r) {
		var t, a, n, i, s, f, l, o, c, u, h, d, v, p;
		var m = sw(e);
		var b = nw(m[1][0].data) >>> 0;
		var g = nw(m[2][0].data) >>> 0;
		var w = ((a = (t = m[8]) == null ? void 0 : t[0]) == null ? void 0 : a.data) && nw(m[8][0].data) > 0 || false;
		var k, T;
		if (((i = (n = m[7]) == null ? void 0 : n[0]) == null ? void 0 : i.data) && r != 0) {
			k = (f = (s = m[7]) == null ? void 0 : s[0]) == null ? void 0 : f.data;
			T = (o = (l = m[6]) == null ? void 0 : l[0]) == null ? void 0 : o.data
		} else if (((u = (c = m[4]) == null ? void 0 : c[0]) == null ? void 0 : u.data) && r != 1) {
			k = (d = (h = m[4]) == null ? void 0 : h[0]) == null ? void 0 : d.data;
			T = (p = (v = m[3]) == null ? void 0 : v[0]) == null ? void 0 : p.data
		} else throw "NUMBERS Tile missing ".concat(r, " cell storage");
		var y = w ? 4 : 1;
		var E = jg(k);
		var _ = [];
		for (var S = 0; S < k.length / 2; ++S) {
			var x = E.getUint16(S * 2, true);
			if (x < 65535) _.push([S, x])
		}
		if (_.length != g) throw "Expected ".concat(g, " cells, found ").concat(_.length);
		var A = [];
		for (S = 0; S < _.length - 1; ++S) A[_[S][0]] = T[Gg](_[S][1] * y, _[S + 1][1] * y);
		if (_.length >= 1) A[_[_.length - 1][0]] = T[Gg](_[_.length - 1][1] * y);
		return {
			R: b,
			cells: A
		}
	}
	function Aw(e, r) {
		var t;
		var a = sw(r.data);
		var n = -1;
		if ((t = a == null ? void 0 : a[7]) == null ? void 0 : t[0]) {
			if (nw(a[7][0].data) >>> 0) n = 1;
			else n = 0
		}
		var i = lw(a[5],
		function(e) {
			return xw(e, n)
		});
		return {
			nrows: nw(a[4][0].data) >>> 0,
			data: i.reduce(function(e, r) {
				if (!e[r.R]) e[r.R] = [];
				r.cells.forEach(function(t, a) {
					if (e[r.R][a]) throw new Error("Duplicate cell r=".concat(r.R, " c=").concat(a));
					e[r.R][a] = t
				});
				return e
			},
			[])
		}
	}
	function Cw(e, r) {
		var t, a, n, i, s, f, l, o, c, u;
		var h = {
			t: "",
			a: ""
		};
		var d = sw(r);
		if ((a = (t = d == null ? void 0 : d[1]) == null ? void 0 : t[0]) == null ? void 0 : a.data) h.t = Kg((i = (n = d == null ? void 0 : d[1]) == null ? void 0 : n[0]) == null ? void 0 : i.data) || "";
		if ((f = (s = d == null ? void 0 : d[3]) == null ? void 0 : s[0]) == null ? void 0 : f.data) {
			var v = e[Tw((o = (l = d == null ? void 0 : d[3]) == null ? void 0 : l[0]) == null ? void 0 : o.data)][0];
			var p = sw(v.data);
			if ((u = (c = p[1]) == null ? void 0 : c[0]) == null ? void 0 : u.data) h.a = Kg(p[1][0].data)
		}
		if (d == null ? void 0 : d[4]) {
			h.replies = [];
			d[4].forEach(function(r) {
				var t = e[Tw(r.data)][0];
				h.replies.push(Cw(e, t.data))
			})
		}
		return h
	}
	function Rw(e) {
		var r = [];
		r.push({
			t: e.t || "",
			a: e.a,
			T: e.replies && e.replies.length > 0
		});
		if (e.replies) e.replies.forEach(function(e) {
			r.push({
				t: e.t || "",
				a: e.a,
				T: true
			})
		});
		return r
	}
	function Ow(e) {
		var r = {
			a: "",
			t: "",
			replies: []
		};
		for (var t = 0; t < e.length; ++t) {
			if (t == 0) {
				r.a = e[t].a;
				r.t = e[t].t
			} else {
				r.replies.push({
					a: e[t].a,
					t: e[t].t
				})
			}
		}
		return r
	}
	function Iw(e, r, t, a) {
		var n, i, s, f, l, o, c, u, h, d;
		var v = sw(r.data);
		var p = {
			s: {
				r: 0,
				c: 0
			},
			e: {
				r: 0,
				c: 0
			}
		};
		p.e.r = (nw(v[6][0].data) >>> 0) - 1;
		if (p.e.r < 0) throw new Error("Invalid row varint ".concat(v[6][0].data));
		p.e.c = (nw(v[7][0].data) >>> 0) - 1;
		if (p.e.c < 0) throw new Error("Invalid col varint ".concat(v[7][0].data));
		t["!ref"] = Va(p);
		var m = t["!data"] != null,
		b = t;
		var g = sw(v[4][0].data);
		var w = vw();
		if ((n = g[4]) == null ? void 0 : n[0]) w.sst = Sw(e, e[Tw(g[4][0].data)][0]);
		if ((i = g[6]) == null ? void 0 : i[0]) w.fmla = Sw(e, e[Tw(g[6][0].data)][0]);
		if ((s = g[11]) == null ? void 0 : s[0]) w.ofmt = Sw(e, e[Tw(g[11][0].data)][0]);
		if ((f = g[12]) == null ? void 0 : f[0]) w.ferr = Sw(e, e[Tw(g[12][0].data)][0]);
		if ((l = g[17]) == null ? void 0 : l[0]) w.rsst = Sw(e, e[Tw(g[17][0].data)][0]);
		if ((o = g[19]) == null ? void 0 : o[0]) w.cmnt = Sw(e, e[Tw(g[19][0].data)][0]);
		if ((c = g[22]) == null ? void 0 : c[0]) w.nfmt = Sw(e, e[Tw(g[22][0].data)][0]);
		var k = sw(g[3][0].data);
		var T = 0;
		if (! ((u = g[9]) == null ? void 0 : u[0])) throw "NUMBERS file missing row tree";
		var y = sw(g[9][0].data)[1].map(function(e) {
			return sw(e.data)
		});
		y.forEach(function(r) {
			T = nw(r[1][0].data);
			var n = nw(r[2][0].data);
			var i = k[1][n];
			if (!i) throw "NUMBERS missing tile " + n;
			var s = sw(i.data);
			var f = e[Tw(s[2][0].data)][0];
			var l = nw(f.meta[1][0].data);
			if (l != 6002) throw new Error("6001 unexpected reference to ".concat(l));
			var o = Aw(e, f);
			o.data.forEach(function(e, r) {
				e.forEach(function(e, n) {
					var i = kw(e, w, a);
					if (i) {
						if (m) {
							if (!b["!data"][T + r]) b["!data"][T + r] = [];
							b["!data"][T + r][n] = i
						} else {
							t[La(n) + Na(T + r)] = i
						}
					}
				})
			});
			T += o.nrows
		});
		if ((h = g[13]) == null ? void 0 : h[0]) {
			var E = e[Tw(g[13][0].data)][0];
			var _ = nw(E.meta[1][0].data);
			if (_ != 6144) throw new Error("Expected merge type 6144, found ".concat(_));
			t["!merges"] = (d = sw(E.data)) == null ? void 0 : d[1].map(function(e) {
				var r = sw(e.data);
				var t = jg(sw(r[1][0].data)[1][0].data),
				a = jg(sw(r[2][0].data)[1][0].data);
				return {
					s: {
						r: t.getUint16(0, true),
						c: t.getUint16(2, true)
					},
					e: {
						r: t.getUint16(0, true) + a.getUint16(0, true) - 1,
						c: t.getUint16(2, true) + a.getUint16(2, true) - 1
					}
				}
			})
		}
	}
	function Nw(e, r, t) {
		var a = sw(r.data);
		var n = {
			"!ref": "A1"
		};
		if (t == null ? void 0 : t.dense) n["!data"] = [];
		var i = e[Tw(a[2][0].data)];
		var s = nw(i[0].meta[1][0].data);
		if (s != 6001) throw new Error("6000 unexpected reference to ".concat(s));
		Iw(e, i[0], n, t);
		return n
	}
	function Fw(e, r, t) {
		var a;
		var n = sw(r.data);
		var i = {
			name: ((a = n[1]) == null ? void 0 : a[0]) ? Kg(n[1][0].data) : "",
			sheets: []
		};
		var s = lw(n[2], Tw);
		s.forEach(function(r) {
			e[r].forEach(function(r) {
				var a = nw(r.meta[1][0].data);
				if (a == 6e3) i.sheets.push(Nw(e, r, t))
			})
		});
		return i
	}
	function Dw(e, r, t) {
		var a;
		var n = jk();
		n.Workbook = {
			WBProps: {
				date1904: true
			}
		};
		var i = sw(r.data);
		if ((a = i[2]) == null ? void 0 : a[0]) throw new Error("Keynote presentations are not supported");
		var s = lw(i[1], Tw);
		s.forEach(function(r) {
			e[r].forEach(function(r) {
				var a = nw(r.meta[1][0].data);
				if (a == 2) {
					var i = Fw(e, r, t);
					i.sheets.forEach(function(e, r) {
						Kk(n, e, r == 0 ? i.name: i.name + "_" + r, true)
					})
				}
			})
		});
		if (n.SheetNames.length == 0) throw new Error("Empty NUMBERS file");
		n.bookType = "numbers";
		return n
	}
	function Pw(e, r) {
		var t, a, n, i, s, f, l;
		var o = {},
		c = [];
		e.FullPaths.forEach(function(e) {
			if (e.match(/\.iwpv2/)) throw new Error("Unsupported password protection")
		});
		e.FileIndex.forEach(function(e) {
			if (!e.name.match(/\.iwa$/)) return;
			if (e.content[0] != 0) return;
			var r;
			try {
				r = hw(e.content)
			} catch(t) {
				return console.log("?? " + e.content.length + " " + (t.message || t))
			}
			var a;
			try {
				a = ow(r)
			} catch(t) {
				return console.log("## " + (t.message || t))
			}
			a.forEach(function(e) {
				o[e.id] = e.messages;
				c.push(e.id)
			})
		});
		if (!c.length) throw new Error("File has no messages");
		if (((n = (a = (t = o == null ? void 0 : o[1]) == null ? void 0 : t[0].meta) == null ? void 0 : a[1]) == null ? void 0 : n[0].data) && nw(o[1][0].meta[1][0].data) == 1e4) throw new Error("Pages documents are not supported");
		var u = ((l = (f = (s = (i = o == null ? void 0 : o[1]) == null ? void 0 : i[0]) == null ? void 0 : s.meta) == null ? void 0 : f[1]) == null ? void 0 : l[0].data) && nw(o[1][0].meta[1][0].data) == 1 && o[1][0];
		if (!u) c.forEach(function(e) {
			o[e].forEach(function(e) {
				var r = nw(e.meta[1][0].data) >>> 0;
				if (r == 1) {
					if (!u) u = e;
					else throw new Error("Document has multiple roots")
				}
			})
		});
		if (!u) throw new Error("Cannot find Document root");
		return Dw(o, u, r)
	}
	function Lw(e, r, t) {
		var a, n, i;
		var s = [[], [{
			type: 0,
			data: rw(0)
		}], [{
			type: 0,
			data: rw(0)
		}], [{
			type: 2,
			data: new Uint8Array([])
		}], [{
			type: 2,
			data: new Uint8Array(Array.from({
				length: 510
			},
			function() {
				return 255
			}))
		}], [{
			type: 0,
			data: rw(5)
		}], [{
			type: 2,
			data: new Uint8Array([])
		}], [{
			type: 2,
			data: new Uint8Array(Array.from({
				length: 510
			},
			function() {
				return 255
			}))
		}], [{
			type: 0,
			data: rw(1)
		}]];
		if (! ((a = s[6]) == null ? void 0 : a[0]) || !((n = s[7]) == null ? void 0 : n[0])) throw "Mutation only works on post-BNC storages!";
		var f = 0;
		if (s[7][0].data.length < 2 * e.length) {
			var l = new Uint8Array(2 * e.length);
			l.set(s[7][0].data);
			s[7][0].data = l
		}
		if (s[4][0].data.length < 2 * e.length) {
			var o = new Uint8Array(2 * e.length);
			o.set(s[4][0].data);
			s[4][0].data = o
		}
		var c = jg(s[7][0].data),
		u = 0,
		h = [];
		var d = jg(s[4][0].data),
		v = 0,
		p = [];
		var m = t ? 4 : 1;
		for (var b = 0; b < e.length; ++b) {
			if (e[b] == null || e[b].t == "z" && !((i = e[b].c) == null ? void 0 : i.length) || e[b].t == "e") {
				c.setUint16(b * 2, 65535, true);
				d.setUint16(b * 2, 65535);
				continue
			}
			c.setUint16(b * 2, u / m, true);
			d.setUint16(b * 2, v / m, true);
			var g, w;
			switch (e[b].t) {
			case "d":
				if (e[b].v instanceof Date) {
					g = gw(e[b], r);
					w = ww(e[b], r);
					break
				}
				g = gw(e[b], r);
				w = ww(e[b], r);
				break;
			case "s":
				;
			case "n":
				;
			case "b":
				;
			case "z":
				g = gw(e[b], r);
				w = ww(e[b], r);
				break;
			default:
				throw new Error("Unsupported value " + e[b]);
			}
			h.push(g);
			u += g.length; {
				p.push(w);
				v += w.length
			}++f
		}
		s[2][0].data = rw(f);
		s[5][0].data = rw(5);
		for (; b < s[7][0].data.length / 2; ++b) {
			c.setUint16(b * 2, 65535, true);
			d.setUint16(b * 2, 65535, true)
		}
		s[6][0].data = Zg(h);
		s[3][0].data = Zg(p);
		s[8] = [{
			type: 0,
			data: rw(t ? 1 : 0)
		}];
		return s
	}
	function Mw(e, r) {
		return {
			meta: [[], [{
				type: 0,
				data: rw(e)
			}]],
			data: r
		}
	}
	function Uw(e, r) {
		if (!r.last) r.last = 927262;
		for (var t = r.last; t < 2e6; ++t) if (!r[t]) {
			r[r.last = t] = e;
			return t
		}
		throw new Error("Too many messages")
	}
	function Bw(e) {
		var r = {};
		var t = [];
		e.FileIndex.map(function(r, t) {
			return [r, e.FullPaths[t]]
		}).forEach(function(e) {
			var a = e[0],
			n = e[1];
			if (a.type != 2) return;
			if (!a.name.match(/\.iwa/)) return;
			if (a.content[0] != 0) return;
			ow(hw(a.content)).forEach(function(e) {
				t.push(e.id);
				r[e.id] = {
					deps: [],
					location: n,
					type: nw(e.messages[0].meta[1][0].data)
				}
			})
		});
		e.FileIndex.forEach(function(e) {
			if (!e.name.match(/\.iwa/)) return;
			if (e.content[0] != 0) return;
			ow(hw(e.content)).forEach(function(e) {
				e.messages.forEach(function(t) { [5, 6].forEach(function(a) {
						if (!t.meta[a]) return;
						t.meta[a].forEach(function(t) {
							r[e.id].deps.push(nw(t.data))
						})
					})
				})
			})
		});
		return r
	}
	function Ww(e, r, t) {
		return fw([[], [{
			type: 0,
			data: rw(1)
		}], [], [{
			type: 5,
			data: new Uint8Array(Float32Array.from([e / 255]).buffer)
		}], [{
			type: 5,
			data: new Uint8Array(Float32Array.from([r / 255]).buffer)
		}], [{
			type: 5,
			data: new Uint8Array(Float32Array.from([t / 255]).buffer)
		}], [{
			type: 5,
			data: new Uint8Array(Float32Array.from([1]).buffer)
		}], [], [], [], [], [], [{
			type: 0,
			data: rw(1)
		}]])
	}
	function zw(e) {
		switch (e) {
		case 0:
			return Ww(99, 222, 171);
		case 1:
			return Ww(162, 197, 240);
		case 2:
			return Ww(255, 189, 189);
		}
		return Ww(Math.random() * 255, Math.random() * 255, Math.random() * 255)
	}
	function Hw(e, r) {
		if (!r || !r.numbers) throw new Error("Must pass a `numbers` option -- check the README");
		var t = Qe.read(r.numbers, {
			type: "base64"
		});
		var a = Bw(t);
		var n = $w(t, a, 1);
		if (n == null) throw "Could not find message ".concat(1, " in Numbers template");
		var i = lw(sw(n.messages[0].data)[1], Tw);
		if (i.length > 1) throw new Error("Template NUMBERS file must have exactly one sheet");
		e.SheetNames.forEach(function(r, s) {
			if (s >= 1) {
				Yw(t, a, s + 1);
				n = $w(t, a, 1);
				i = lw(sw(n.messages[0].data)[1], Tw)
			}
			Zw(t, a, e.Sheets[r], r, s, i[s])
		});
		return t
	}
	function Vw(e, r, t, a) {
		var n = Qe.find(e, r[t].location);
		if (!n) throw "Could not find ".concat(r[t].location, " in Numbers template");
		var i = ow(hw(n.content));
		var s = i.find(function(e) {
			return e.id == t
		});
		a(s, i);
		n.content = dw(cw(i));
		n.size = n.content.length
	}
	function $w(e, r, t) {
		var a = Qe.find(e, r[t].location);
		if (!a) throw "Could not find ".concat(r[t].location, " in Numbers template");
		var n = ow(hw(a.content));
		var i = n.find(function(e) {
			return e.id == t
		});
		return i
	}
	function Xw(e, r, t) {
		e[3].push({
			type: 2,
			data: fw([[], [{
				type: 0,
				data: rw(r)
			}], [{
				type: 2,
				data: Yg(t.replace(/-.*$/, ""))
			}], [{
				type: 2,
				data: Yg(t)
			}], [{
				type: 2,
				data: new Uint8Array([2, 0, 0])
			}], [{
				type: 2,
				data: new Uint8Array([2, 0, 0])
			}], [], [], [], [], [{
				type: 0,
				data: rw(0)
			}], [], [{
				type: 0,
				data: rw(0)
			}]])
		});
		e[1] = [{
			type: 0,
			data: rw(Math.max(r + 1, nw(e[1][0].data)))
		}]
	}
	function Gw(e, r, t, a, n, i) {
		if (!i) i = Uw({
			deps: [],
			location: "",
			type: r
		},
		n);
		var s = "".concat(a, "-").concat(i, ".iwa");
		n[i].location = "Root Entry" + s;
		Qe.utils.cfb_add(e, s, dw(cw([{
			id: i,
			messages: [Mw(r, fw(t))]
		}])));
		var f = s.replace(/^[\/]/, "").replace(/^Index\//, "").replace(/\.iwa$/, "");
		Vw(e, n, 2,
		function(e) {
			var r = sw(e.messages[0].data);
			Xw(r, i || 0, f);
			e.messages[0].data = fw(r)
		});
		return i
	}
	function jw(e, r, t, a) {
		var n = r[t].location.replace(/^Root Entry\//, "").replace(/^Index\//, "").replace(/\.iwa$/, "");
		var i = e[3].findIndex(function(e) {
			var r, t;
			var a = sw(e.data);
			if ((r = a[3]) == null ? void 0 : r[0]) return Kg(a[3][0].data) == n;
			if (((t = a[2]) == null ? void 0 : t[0]) && Kg(a[2][0].data) == n) return true;
			return false
		});
		var s = sw(e[3][i].data);
		if (!s[6]) s[6] = []; (Array.isArray(a) ? a: [a]).forEach(function(e) {
			s[6].push({
				type: 2,
				data: fw([[], [{
					type: 0,
					data: rw(e)
				}]])
			})
		});
		e[3][i].data = fw(s)
	}
	function Kw(e, r, t, a) {
		var n = r[t].location.replace(/^Root Entry\//, "").replace(/^Index\//, "").replace(/\.iwa$/, "");
		var i = e[3].findIndex(function(e) {
			var r, t;
			var a = sw(e.data);
			if ((r = a[3]) == null ? void 0 : r[0]) return Kg(a[3][0].data) == n;
			if (((t = a[2]) == null ? void 0 : t[0]) && Kg(a[2][0].data) == n) return true;
			return false
		});
		var s = sw(e[3][i].data);
		if (!s[6]) s[6] = [];
		s[6] = s[6].filter(function(e) {
			return nw(sw(e.data)[1][0].data) != a
		});
		e[3][i].data = fw(s)
	}
	function Yw(e, r, t) {
		var a = -1,
		n = -1;
		var i = {};
		Vw(e, r, 1,
		function(t, s) {
			var f = sw(t.messages[0].data);
			a = Tw(sw(t.messages[0].data)[1][0].data);
			n = Uw({
				deps: [1],
				location: r[a].location,
				type: 2
			},
			r);
			i[a] = n;
			Ew(t, n);
			f[1].push({
				type: 2,
				data: yw(n)
			});
			var l = $w(e, r, a);
			l.id = n;
			if (r[1].location == r[n].location) s.push(l);
			else Vw(e, r, n,
			function(e, r) {
				return r.push(l)
			});
			t.messages[0].data = fw(f)
		});
		var s = -1;
		Vw(e, r, n,
		function(t, a) {
			var f = sw(t.messages[0].data);
			for (var l = 3; l <= 69; ++l) delete f[l];
			var o = lw(f[2], Tw);
			o.forEach(function(e) {
				return _w(t, e)
			});
			s = Uw({
				deps: [n],
				location: r[o[0]].location,
				type: r[o[0]].type
			},
			r);
			Ew(t, s);
			i[o[0]] = s;
			f[2] = [{
				type: 2,
				data: yw(s)
			}];
			var c = $w(e, r, o[0]);
			c.id = s;
			if (r[o[0]].location == r[n].location) a.push(c);
			else {
				Vw(e, r, 2,
				function(e) {
					var t = sw(e.messages[0].data);
					jw(t, r, n, s);
					e.messages[0].data = fw(t)
				});
				Vw(e, r, s,
				function(e, r) {
					return r.push(c)
				})
			}
			t.messages[0].data = fw(f)
		});
		var f = -1;
		Vw(e, r, s,
		function(t, a) {
			var n = sw(t.messages[0].data);
			var l = sw(n[1][0].data);
			for (var o = 3; o <= 69; ++o) delete l[o];
			var c = Tw(l[2][0].data);
			l[2][0].data = yw(i[c]);
			n[1][0].data = fw(l);
			var u = Tw(n[2][0].data);
			_w(t, u);
			f = Uw({
				deps: [s],
				location: r[u].location,
				type: r[u].type
			},
			r);
			Ew(t, f);
			i[u] = f;
			n[2][0].data = yw(f);
			var h = $w(e, r, u);
			h.id = f;
			if (r[s].location == r[f].location) a.push(h);
			else Vw(e, r, f,
			function(e, r) {
				return r.push(h)
			});
			t.messages[0].data = fw(n)
		});
		Vw(e, r, f,
		function(a, n) {
			var s, l;
			var o = sw(a.messages[0].data);
			var c = Kg(o[1][0].data),
			u = c.replace(/-[A-Z0-9]*/, "-".concat(("0000" + t.toString(16)).slice(-4)));
			o[1][0].data = Yg(u); [12, 13, 29, 31, 32, 33, 39, 44, 47, 81, 82, 84].forEach(function(e) {
				return delete o[e]
			});
			if (o[45]) {
				var h = sw(o[45][0].data);
				var d = Tw(h[1][0].data);
				_w(a, d);
				delete o[45]
			}
			if (o[70]) {
				var v = sw(o[70][0].data); (s = v[2]) == null ? void 0 : s.forEach(function(e) {
					var r = sw(e.data); [2, 3].map(function(e) {
						return r[e][0]
					}).forEach(function(e) {
						var r = sw(e.data);
						if (!r[8]) return;
						var t = Tw(r[8][0].data);
						_w(a, t)
					})
				});
				delete o[70]
			} [46, 30, 34, 35, 36, 38, 48, 49, 60, 61, 62, 63, 64, 71, 72, 73, 74, 75, 85, 86, 87, 88, 89].forEach(function(e) {
				if (!o[e]) return;
				var r = Tw(o[e][0].data);
				delete o[e];
				_w(a, r)
			});
			var p = sw(o[4][0].data); { [2, 4, 5, 6, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22].forEach(function(t) {
					var s;
					if (! ((s = p[t]) == null ? void 0 : s[0])) return;
					var l = Tw(p[t][0].data);
					var o = Uw({
						deps: [f],
						location: r[l].location,
						type: r[l].type
					},
					r);
					_w(a, l);
					Ew(a, o);
					i[l] = o;
					var c = $w(e, r, l);
					c.id = o;
					if (r[l].location == r[f].location) n.push(c);
					else {
						r[o].location = r[l].location.replace(l.toString(), o.toString());
						if (r[o].location == r[l].location) r[o].location = r[o].location.replace(/\.iwa/, "-".concat(o, ".iwa"));
						Qe.utils.cfb_add(e, r[o].location, dw(cw([c])));
						var u = r[o].location.replace(/^Root Entry\//, "").replace(/^Index\//, "").replace(/\.iwa$/, "");
						Vw(e, r, 2,
						function(e) {
							var t = sw(e.messages[0].data);
							Xw(t, o, u);
							jw(t, r, f, o);
							e.messages[0].data = fw(t)
						})
					}
					p[t][0].data = yw(o)
				});
				var m = sw(p[1][0].data); { (l = m[2]) == null ? void 0 : l.forEach(function(t) {
						var s = Tw(t.data);
						var l = Uw({
							deps: [f],
							location: r[s].location,
							type: r[s].type
						},
						r);
						_w(a, s);
						Ew(a, l);
						i[s] = l;
						var o = $w(e, r, s);
						o.id = l;
						if (r[s].location == r[f].location) {
							n.push(o)
						} else {
							r[l].location = r[s].location.replace(s.toString(), l.toString());
							if (r[l].location == r[s].location) r[l].location = r[l].location.replace(/\.iwa/, "-".concat(l, ".iwa"));
							Qe.utils.cfb_add(e, r[l].location, dw(cw([o])));
							var c = r[l].location.replace(/^Root Entry\//, "").replace(/^Index\//, "").replace(/\.iwa$/, "");
							Vw(e, r, 2,
							function(e) {
								var t = sw(e.messages[0].data);
								Xw(t, l, c);
								jw(t, r, f, l);
								e.messages[0].data = fw(t)
							})
						}
						t.data = yw(l)
					})
				}
				p[1][0].data = fw(m);
				var b = sw(p[3][0].data); {
					b[1].forEach(function(t) {
						var n = sw(t.data);
						var s = Tw(n[2][0].data);
						var l = i[s];
						if (!i[s]) {
							l = Uw({
								deps: [f],
								location: "",
								type: r[s].type
							},
							r);
							r[l].location = "Root Entry/Index/Tables/Tile-".concat(l, ".iwa");
							i[s] = l;
							var o = $w(e, r, s);
							o.id = l;
							_w(a, s);
							Ew(a, l);
							Qe.utils.cfb_add(e, "/Index/Tables/Tile-".concat(l, ".iwa"), dw(cw([o])));
							Vw(e, r, 2,
							function(e) {
								var t = sw(e.messages[0].data);
								t[3].push({
									type: 2,
									data: fw([[], [{
										type: 0,
										data: rw(l)
									}], [{
										type: 2,
										data: Yg("Tables/Tile")
									}], [{
										type: 2,
										data: Yg("Tables/Tile-".concat(l))
									}], [{
										type: 2,
										data: new Uint8Array([2, 0, 0])
									}], [{
										type: 2,
										data: new Uint8Array([2, 0, 0])
									}], [], [], [], [], [{
										type: 0,
										data: rw(0)
									}], [], [{
										type: 0,
										data: rw(0)
									}]])
								});
								t[1] = [{
									type: 0,
									data: rw(Math.max(l + 1, nw(t[1][0].data)))
								}];
								jw(t, r, f, l);
								e.messages[0].data = fw(t)
							})
						}
						n[2][0].data = yw(l);
						t.data = fw(n)
					})
				}
				p[3][0].data = fw(b)
			}
			o[4][0].data = fw(p);
			a.messages[0].data = fw(o)
		})
	}
	function Zw(e, r, t, a, n, i) {
		var s = [];
		Vw(e, r, i,
		function(e) {
			var r = sw(e.messages[0].data); {
				r[1] = [{
					type: 2,
					data: Yg(a)
				}];
				s = lw(r[2], Tw)
			}
			e.messages[0].data = fw(r)
		});
		var f = $w(e, r, s[0]);
		var l = Tw(sw(f.messages[0].data)[2][0].data);
		Vw(e, r, l,
		function(a, n) {
			return qw(e, r, t, a, n, l)
		})
	}
	var Jw = true;
	function qw(e, r, t, a, n, i) {
		if (!t["!ref"]) throw new Error("Cannot export empty sheet to NUMBERS");
		var s = Ha(t["!ref"]);
		s.s.r = s.s.c = 0;
		var f = false;
		if (s.e.c > 999) {
			f = true;
			s.e.c = 999
		}
		if (s.e.r > 999999) {
			f = true;
			s.e.r = 999999
		}
		if (f) console.error("Truncating to ".concat(Va(s)));
		var l = [];
		if (t["!data"]) l = t["!data"];
		else {
			var o = [];
			for (var c = 0; c <= s.e.c; ++c) o[c] = La(c);
			for (var u = 0; u <= s.e.r; ++u) {
				l[u] = [];
				var h = "" + (u + 1);
				for (c = 0; c <= s.e.c; ++c) {
					var d = t[o[c] + h];
					if (!d) continue;
					l[u][c] = d
				}
			}
		}
		var v = {
			cmnt: [{
				a: "~54ee77S~",
				t: "... the people who are crazy enough to think they can change the world, are the ones who do."
			}],
			ferr: [],
			fmla: [],
			nfmt: [],
			ofmt: [],
			rsst: [{
				v: "~54ee77S~",
				l: "https://sheetjs.com/"
			}],
			sst: ["~Sh33tJ5~"]
		};
		var p = sw(a.messages[0].data); {
			p[6][0].data = rw(s.e.r + 1);
			p[7][0].data = rw(s.e.c + 1);
			delete p[46];
			var m = sw(p[4][0].data); {
				var b = Tw(sw(m[1][0].data)[2][0].data);
				Vw(e, r, b,
				function(e, r) {
					var t;
					var a = sw(e.messages[0].data);
					if ((t = a == null ? void 0 : a[2]) == null ? void 0 : t[0]) for (var n = 0; n < l.length; ++n) {
						var i = sw(a[2][0].data);
						i[1][0].data = rw(n);
						i[4][0].data = rw(l[n].length);
						a[2][n] = {
							type: a[2][0].type,
							data: fw(i)
						}
					}
					e.messages[0].data = fw(a)
				});
				var g = Tw(m[2][0].data);
				Vw(e, r, g,
				function(e, r) {
					var t = sw(e.messages[0].data);
					for (var a = 0; a <= s.e.c; ++a) {
						var n = sw(t[2][0].data);
						n[1][0].data = rw(a);
						n[4][0].data = rw(s.e.r + 1);
						t[2][a] = {
							type: t[2][0].type,
							data: fw(n)
						}
					}
					e.messages[0].data = fw(t)
				});
				var w = sw(m[9][0].data);
				w[1] = [];
				var k = sw(m[3][0].data); {
					var T = 256;
					k[2] = [{
						type: 0,
						data: rw(T)
					}];
					var y = Tw(sw(k[1][0].data)[2][0].data);
					var E = function() {
						var t = $w(e, r, 2);
						var a = sw(t.messages[0].data);
						var n = a[3].filter(function(e) {
							return nw(sw(e.data)[1][0].data) == y
						});
						return (n == null ? void 0 : n.length) ? nw(sw(n[0].data)[12][0].data) : 0
					} (); {
						Qe.utils.cfb_del(e, r[y].location);
						Vw(e, r, 2,
						function(e) {
							var t = sw(e.messages[0].data);
							t[3] = t[3].filter(function(e) {
								return nw(sw(e.data)[1][0].data) != y
							});
							Kw(t, r, i, y);
							e.messages[0].data = fw(t)
						});
						_w(a, y)
					}
					k[1] = [];
					var _ = Math.ceil((s.e.r + 1) / T);
					for (var S = 0; S < _; ++S) {
						var x = Uw({
							deps: [],
							location: "",
							type: 6002
						},
						r);
						r[x].location = "Root Entry/Index/Tables/Tile-".concat(x, ".iwa");
						var A = [[], [{
							type: 0,
							data: rw(0)
						}], [{
							type: 0,
							data: rw(Math.min(s.e.r + 1, (S + 1) * T))
						}], [{
							type: 0,
							data: rw(0)
						}], [{
							type: 0,
							data: rw(Math.min((S + 1) * T, s.e.r + 1) - S * T)
						}], [], [{
							type: 0,
							data: rw(5)
						}], [{
							type: 0,
							data: rw(1)
						}], [{
							type: 0,
							data: rw(Jw ? 1 : 0)
						}]];
						for (var C = S * T; C <= Math.min(s.e.r, (S + 1) * T - 1); ++C) {
							var R = Lw(l[C], v, Jw);
							R[1][0].data = rw(C - S * T);
							A[5].push({
								data: fw(R),
								type: 2
							})
						}
						k[1].push({
							type: 2,
							data: fw([[], [{
								type: 0,
								data: rw(S)
							}], [{
								type: 2,
								data: yw(x)
							}]])
						});
						var O = {
							id: x,
							messages: [Mw(6002, fw(A))]
						};
						var I = dw(cw([O]));
						Qe.utils.cfb_add(e, "/Index/Tables/Tile-".concat(x, ".iwa"), I);
						Vw(e, r, 2,
						function(e) {
							var t = sw(e.messages[0].data);
							t[3].push({
								type: 2,
								data: fw([[], [{
									type: 0,
									data: rw(x)
								}], [{
									type: 2,
									data: Yg("Tables/Tile")
								}], [{
									type: 2,
									data: Yg("Tables/Tile-".concat(x))
								}], [{
									type: 2,
									data: new Uint8Array([2, 0, 0])
								}], [{
									type: 2,
									data: new Uint8Array([2, 0, 0])
								}], [], [], [], [], [{
									type: 0,
									data: rw(0)
								}], [], [{
									type: 0,
									data: rw(E)
								}]])
							});
							t[1] = [{
								type: 0,
								data: rw(Math.max(x + 1, nw(t[1][0].data)))
							}];
							jw(t, r, i, x);
							e.messages[0].data = fw(t)
						});
						Ew(a, x);
						w[1].push({
							type: 2,
							data: fw([[], [{
								type: 0,
								data: rw(S * T)
							}], [{
								type: 0,
								data: rw(S)
							}]])
						})
					}
				}
				m[3][0].data = fw(k);
				m[9][0].data = fw(w);
				m[10] = [{
					type: 2,
					data: new Uint8Array([])
				}];
				if (t["!merges"]) {
					var N = Uw({
						type: 6144,
						deps: [i],
						location: r[i].location
					},
					r);
					n.push({
						id: N,
						messages: [Mw(6144, fw([[], t["!merges"].map(function(e) {
							return {
								type: 2,
								data: fw([[], [{
									type: 2,
									data: fw([[], [{
										type: 5,
										data: new Uint8Array(new Uint16Array([e.s.r, e.s.c]).buffer)
									}]])
								}], [{
									type: 2,
									data: fw([[], [{
										type: 5,
										data: new Uint8Array(new Uint16Array([e.e.r - e.s.r + 1, e.e.c - e.s.c + 1]).buffer)
									}]])
								}]])
							}
						})]))]
					});
					m[13] = [{
						type: 2,
						data: yw(N)
					}];
					Vw(e, r, 2,
					function(e) {
						var t = sw(e.messages[0].data);
						jw(t, r, i, N);
						e.messages[0].data = fw(t)
					});
					Ew(a, N)
				} else delete m[13];
				var F = Tw(m[4][0].data);
				Vw(e, r, F,
				function(e) {
					var r = sw(e.messages[0].data); {
						r[3] = [];
						v.sst.forEach(function(e, t) {
							if (t == 0) return;
							r[3].push({
								type: 2,
								data: fw([[], [{
									type: 0,
									data: rw(t)
								}], [{
									type: 0,
									data: rw(1)
								}], [{
									type: 2,
									data: Yg(e)
								}]])
							})
						})
					}
					e.messages[0].data = fw(r)
				});
				var D = Tw(m[17][0].data);
				Vw(e, r, D,
				function(t) {
					var a = sw(t.messages[0].data);
					a[3] = [];
					var n = [904980, 903835, 903815, 903845];
					v.rsst.forEach(function(i, s) {
						if (s == 0) return;
						var f = [[], [{
							type: 0,
							data: new Uint8Array([5])
						}], [], [{
							type: 2,
							data: Yg(i.v)
						}]];
						f[10] = [{
							type: 0,
							data: new Uint8Array([1])
						}];
						f[19] = [{
							type: 2,
							data: new Uint8Array([10, 6, 8, 0, 18, 2, 101, 110])
						}];
						f[5] = [{
							type: 2,
							data: new Uint8Array([10, 8, 8, 0, 18, 4, 8, 155, 149, 55])
						}];
						f[2] = [{
							type: 2,
							data: new Uint8Array([8, 148, 158, 55])
						}];
						f[6] = [{
							type: 2,
							data: new Uint8Array([10, 6, 8, 0, 16, 0, 24, 0])
						}];
						f[7] = [{
							type: 2,
							data: new Uint8Array([10, 8, 8, 0, 18, 4, 8, 135, 149, 55])
						}];
						f[8] = [{
							type: 2,
							data: new Uint8Array([10, 8, 8, 0, 18, 4, 8, 165, 149, 55])
						}];
						f[14] = [{
							type: 2,
							data: new Uint8Array([10, 6, 8, 0, 16, 0, 24, 0])
						}];
						f[24] = [{
							type: 2,
							data: new Uint8Array([10, 6, 8, 0, 16, 0, 24, 0])
						}];
						var l = Uw({
							deps: [],
							location: "",
							type: 2001
						},
						r);
						var o = [];
						if (i.l) {
							var c = Gw(e, 2032, [[], [], [{
								type: 2,
								data: Yg(i.l)
							}]], "/Index/Tables/DataList", r);
							f[11] = [];
							var u = [[], []];
							if (!u[1]) u[1] = [];
							u[1].push({
								type: 2,
								data: fw([[], [{
									type: 0,
									data: rw(0)
								}], [{
									type: 2,
									data: yw(c)
								}]])
							});
							f[11][0] = {
								type: 2,
								data: fw(u)
							};
							o.push(c)
						}
						Gw(e, 2001, f, "/Index/Tables/DataList", r, l);
						Vw(e, r, l,
						function(e) {
							n.forEach(function(r) {
								return Ew(e, r)
							});
							o.forEach(function(r) {
								return Ew(e, r)
							})
						});
						var h = Gw(e, 6218, [[], [{
							type: 2,
							data: yw(l)
						}], [], [{
							type: 2,
							data: new Uint8Array([13, 255, 255, 255, 0, 18, 10, 16, 255, 255, 1, 24, 255, 255, 255, 255, 7])
						}]], "/Index/Tables/DataList", r);
						Vw(e, r, h,
						function(e) {
							return Ew(e, l)
						});
						a[3].push({
							type: 2,
							data: fw([[], [{
								type: 0,
								data: rw(s)
							}], [{
								type: 0,
								data: rw(1)
							}], [], [], [], [], [], [], [{
								type: 2,
								data: yw(h)
							}]])
						});
						Ew(t, h);
						Vw(e, r, 2,
						function(e) {
							var t = sw(e.messages[0].data);
							jw(t, r, D, h);
							jw(t, r, h, l);
							jw(t, r, l, o);
							jw(t, r, l, n);
							e.messages[0].data = fw(t)
						})
					});
					t.messages[0].data = fw(a)
				});
				if (v.cmnt.length > 1) {
					var P = Tw(m[19][0].data);
					var L = {},
					M = 0;
					Vw(e, r, P,
					function(t) {
						var a = sw(t.messages[0].data); {
							a[3] = [];
							v.cmnt.forEach(function(n, i) {
								if (i == 0) return;
								var s = [];
								if (n.replies) n.replies.forEach(function(t) {
									if (!L[t.a || ""]) L[t.a || ""] = Gw(e, 212, [[], [{
										type: 2,
										data: Yg(t.a || "")
									}], [{
										type: 2,
										data: zw(++M)
									}], [], [{
										type: 0,
										data: rw(0)
									}]], "/Index/Tables/DataList", r);
									var a = L[t.a || ""];
									var n = Gw(e, 3056, [[], [{
										type: 2,
										data: Yg(t.t || "")
									}], [{
										type: 2,
										data: fw([[], [{
											type: 1,
											data: new Uint8Array([0, 0, 0, 128, 116, 109, 182, 65])
										}]])
									}], [{
										type: 2,
										data: yw(a)
									}]], "/Index/Tables/DataList", r);
									Vw(e, r, n,
									function(e) {
										return Ew(e, a)
									});
									s.push(n);
									Vw(e, r, 2,
									function(e) {
										var t = sw(e.messages[0].data);
										jw(t, r, n, a);
										e.messages[0].data = fw(t)
									})
								});
								if (!L[n.a || ""]) L[n.a || ""] = Gw(e, 212, [[], [{
									type: 2,
									data: Yg(n.a || "")
								}], [{
									type: 2,
									data: zw(++M)
								}], [], [{
									type: 0,
									data: rw(0)
								}]], "/Index/Tables/DataList", r);
								var f = L[n.a || ""];
								var l = Gw(e, 3056, [[], [{
									type: 2,
									data: Yg(n.t || "")
								}], [{
									type: 2,
									data: fw([[], [{
										type: 1,
										data: new Uint8Array([0, 0, 0, 128, 116, 109, 182, 65])
									}]])
								}], [{
									type: 2,
									data: yw(f)
								}], s.map(function(e) {
									return {
										type: 2,
										data: yw(e)
									}
								}), [{
									type: 2,
									data: fw([[], [{
										type: 0,
										data: rw(i)
									}], [{
										type: 0,
										data: rw(0)
									}]])
								}]], "/Index/Tables/DataList", r);
								Vw(e, r, l,
								function(e) {
									Ew(e, f);
									s.forEach(function(r) {
										return Ew(e, r)
									})
								});
								a[3].push({
									type: 2,
									data: fw([[], [{
										type: 0,
										data: rw(i)
									}], [{
										type: 0,
										data: rw(1)
									}], [], [], [], [], [], [], [], [{
										type: 2,
										data: yw(l)
									}]])
								});
								Ew(t, l);
								Vw(e, r, 2,
								function(e) {
									var t = sw(e.messages[0].data);
									jw(t, r, P, l);
									jw(t, r, l, f);
									if (s.length) jw(t, r, l, s);
									e.messages[0].data = fw(t)
								})
							})
						}
						a[2][0].data = rw(v.cmnt.length + 1);
						t.messages[0].data = fw(a)
					})
				}
			}
			p[4][0].data = fw(m)
		}
		a.messages[0].data = fw(p)
	}
	function Qw(e) {
		return function r(t) {
			for (var a = 0; a != e.length; ++a) {
				var n = e[a];
				if (t[n[0]] === undefined) t[n[0]] = n[1];
				if (n[2] === "n") t[n[0]] = Number(t[n[0]])
			}
		}
	}
	function ek(e) {
		Qw([["cellNF", false], ["cellHTML", true], ["cellFormula", true], ["cellStyles", false], ["cellText", true], ["cellDates", false], ["sheetStubs", false], ["sheetRows", 0, "n"], ["bookDeps", false], ["bookSheets", false], ["bookProps", false], ["bookFiles", false], ["bookVBA", false], ["password", ""], ["WTF", false]])(e)
	}
	function rk(e) {
		Qw([["cellDates", false], ["bookSST", false], ["bookType", "xlsx"], ["compression", false], ["WTF", false]])(e)
	}
	function tk(e) {
		if (ci.WS.indexOf(e) > -1) return "sheet";
		if (ci.CS && e == ci.CS) return "chart";
		if (ci.DS && e == ci.DS) return "dialog";
		if (ci.MS && e == ci.MS) return "macro";
		return e && e.length ? e: "sheet"
	}
	function ak(e, r) {
		if (!e) return 0;
		try {
			e = r.map(function a(r) {
				if (!r.id) r.id = r.strRelID;
				return [r.name, e["!id"][r.id].Target, tk(e["!id"][r.id].Type)]
			})
		} catch(t) {
			return null
		}
		return ! e || e.length === 0 ? null: e
	}
	function nk(e, r, t, a, n, i, s, f) {
		if (!e || !e["!legdrawel"]) return;
		var l = jr(e["!legdrawel"].Target, a);
		var o = zr(t, l, true);
		if (o) Iu(kt(o), e, f || [])
	}
	function ik(e, r, t, a, n, i, s, f, l, o, c, u) {
		try {
			i[a] = hi(zr(e, t, true), r);
			var h = Wr(e, r);
			var d;
			switch (f) {
			case "sheet":
				d = ib(h, r, n, l, i[a], o, c, u);
				break;
			case "chart":
				d = sb(h, r, n, l, i[a], o, c, u);
				if (!d || !d["!drawel"]) break;
				var v = jr(d["!drawel"].Target, r);
				var p = ui(v);
				var m = Ru(zr(e, v, true), hi(zr(e, p, true), v));
				var b = jr(m, v);
				var g = ui(b);
				d = Em(zr(e, b, true), b, l, hi(zr(e, g, true), b), o, d);
				break;
			case "macro":
				d = fb(h, r, n, l, i[a], o, c, u);
				break;
			case "dialog":
				d = lb(h, r, n, l, i[a], o, c, u);
				break;
			default:
				throw new Error("Unrecognized sheet type " + f);
			}
			s[a] = d;
			var w = [],
			k = [];
			if (i && i[a]) ir(i[a]).forEach(function(t) {
				var n = "";
				if (i[a][t].Type == ci.CMNT) {
					n = jr(i[a][t].Target, r);
					w = ub(Wr(e, n, true), n, l);
					if (!w || !w.length) return;
					Du(d, w, false)
				}
				if (i[a][t].Type == ci.TCMNT) {
					n = jr(i[a][t].Target, r);
					k = k.concat(Mu(Wr(e, n, true), l))
				}
			});
			if (k && k.length) Du(d, k, true, l.people || []);
			nk(d, f, e, r, n, l, o, w)
		} catch(T) {
			if (l.WTF) throw T
		}
	}
	function sk(e) {
		return e.charAt(0) == "/" ? e.slice(1) : e
	}
	function fk(e, r) {
		$e();
		r = r || {};
		ek(r);
		if (Ur(e, "META-INF/manifest.xml")) return Bg(e, r);
		if (Ur(e, "objectdata.xml")) return Bg(e, r);
		if (Ur(e, "Index/Document.iwa")) {
			if (typeof Uint8Array == "undefined") throw new Error("NUMBERS file parsing requires Uint8Array support");
			if (typeof Pw != "undefined") {
				if (e.FileIndex) return Pw(e, r);
				var t = Qe.utils.cfb_new();
				Vr(e).forEach(function(r) {
					$r(t, r, Hr(e, r))
				});
				return Pw(t, r)
			}
			throw new Error("Unsupported NUMBERS file")
		}
		if (!Ur(e, "[Content_Types].xml")) {
			if (Ur(e, "index.xml.gz")) throw new Error("Unsupported NUMBERS 08 file");
			if (Ur(e, "index.xml")) throw new Error("Unsupported NUMBERS 09 file");
			var a = Qe.find(e, "Index.zip");
			if (a) {
				r = Tr(r);
				delete r.type;
				if (typeof a.content == "string") r.type = "binary";
				if (typeof Bun !== "undefined" && Buffer.isBuffer(a.content)) return wk(new Uint8Array(a.content), r);
				return wk(a.content, r)
			}
			throw new Error("Unsupported ZIP file")
		}
		var n = Vr(e);
		var i = li(zr(e, "[Content_Types].xml"));
		var s = false;
		var f, l;
		if (i.workbooks.length === 0) {
			l = "xl/workbook.xml";
			if (Wr(e, l, true)) i.workbooks.push(l)
		}
		if (i.workbooks.length === 0) {
			l = "xl/workbook.bin";
			if (!Wr(e, l, true)) throw new Error("Could not find workbook");
			i.workbooks.push(l);
			s = true;
		}
		if (i.workbooks[0].slice(-3) == "bin") s = true;
		var o = {};
		var c = {};
		if (!r.bookSheets && !r.bookProps) {
			mv = [];
			if (i.sst) try {
				mv = cb(Wr(e, sk(i.sst)), i.sst, r)
			} catch(u) {
				if (r.WTF) throw u
			}
			if (r.cellStyles && i.themes.length) o = au(zr(e, i.themes[0].replace(/^\//, ""), true) || "", r);
			if (i.style) c = ob(Wr(e, sk(i.style)), i.style, o, r)
		}
		i.links.map(function(t) {
			try {
				var a = hi(zr(e, ui(sk(t))), t);
				return db(Wr(e, sk(t)), a, t, r)
			} catch(n) {}
		});
		var h = nb(Wr(e, sk(i.workbooks[0])), i.workbooks[0], r);
		var d = {},
		v = "";
		if (i.coreprops.length) {
			v = Wr(e, sk(i.coreprops[0]), true);
			if (v) d = _i(v);
			if (i.extprops.length !== 0) {
				v = Wr(e, sk(i.extprops[0]), true);
				if (v) Oi(v, d, r)
			}
		}
		var p = {};
		if (!r.bookSheets || r.bookProps) {
			if (i.custprops.length !== 0) {
				v = zr(e, sk(i.custprops[0]), true);
				if (v) p = Fi(v, r)
			}
		}
		var m = {};
		if (r.bookSheets || r.bookProps) {
			if (h.Sheets) f = h.Sheets.map(function N(e) {
				return e.name
			});
			else if (d.Worksheets && d.SheetNames.length > 0) f = d.SheetNames;
			if (r.bookProps) {
				m.Props = d;
				m.Custprops = p
			}
			if (r.bookSheets && typeof f !== "undefined") m.SheetNames = f;
			if (r.bookSheets ? m.SheetNames: r.bookProps) return m
		}
		f = {};
		var b = {};
		if (r.bookDeps && i.calcchain) b = hb(Wr(e, sk(i.calcchain)), i.calcchain, r);
		var g = 0;
		var w = {};
		var k, T; {
			var y = h.Sheets;
			d.Worksheets = y.length;
			d.SheetNames = [];
			for (var E = 0; E != y.length; ++E) {
				d.SheetNames[E] = y[E].name
			}
		}
		var _ = s ? "bin": "xml";
		var S = i.workbooks[0].lastIndexOf("/");
		var x = (i.workbooks[0].slice(0, S + 1) + "_rels/" + i.workbooks[0].slice(S + 1) + ".rels").replace(/^\//, "");
		if (!Ur(e, x)) x = "xl/_rels/workbook." + _ + ".rels";
		var A = hi(zr(e, x, true), x.replace(/_rels.*/, "s5s"));
		if ((i.metadata || []).length >= 1) {
			r.xlmeta = vb(Wr(e, sk(i.metadata[0])), i.metadata[0], r)
		}
		if ((i.people || []).length >= 1) {
			r.people = Bu(Wr(e, sk(i.people[0])), r)
		}
		if (A) A = ak(A, h.Sheets);
		var C = Wr(e, "xl/worksheets/sheet.xml", true) ? 1 : 0;
		e: for (g = 0; g != d.Worksheets; ++g) {
			var R = "sheet";
			if (A && A[g]) {
				k = "xl/" + A[g][1].replace(/[\/]?xl\//, "");
				if (!Ur(e, k)) k = A[g][1];
				if (!Ur(e, k)) k = x.replace(/_rels\/.*$/, "") + A[g][1];
				R = A[g][2]
			} else {
				k = "xl/worksheets/sheet" + (g + 1 - C) + "." + _;
				k = k.replace(/sheet0\./, "sheet.")
			}
			T = k.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
			if (r && r.sheets != null) switch (typeof r.sheets) {
			case "number":
				if (g != r.sheets) continue e;
				break;
			case "string":
				if (d.SheetNames[g].toLowerCase() != r.sheets.toLowerCase()) continue e;
				break;
			default:
				if (Array.isArray && Array.isArray(r.sheets)) {
					var O = false;
					for (var I = 0; I != r.sheets.length; ++I) {
						if (typeof r.sheets[I] == "number" && r.sheets[I] == g) O = 1;
						if (typeof r.sheets[I] == "string" && r.sheets[I].toLowerCase() == d.SheetNames[g].toLowerCase()) O = 1
					}
					if (!O) continue e
				};
			}
			ik(e, k, T, d.SheetNames[g], g, w, f, R, r, h, o, c)
		}
		m = {
			Directory: i,
			Workbook: h,
			Props: d,
			Custprops: p,
			Deps: b,
			Sheets: f,
			SheetNames: d.SheetNames,
			Strings: mv,
			Styles: c,
			Themes: o,
			SSF: Tr(q)
		};
		if (r && r.bookFiles) {
			if (e.files) {
				m.keys = n;
				m.files = e.files
			} else {
				m.keys = [];
				m.files = {};
				e.FullPaths.forEach(function(r, t) {
					r = r.replace(/^Root Entry[\/]/, "");
					m.keys.push(r);
					m.files[r] = e.FileIndex[t]
				})
			}
		}
		if (r && r.bookVBA) {
			if (i.vba.length > 0) m.vbaraw = Wr(e, sk(i.vba[0]), true);
			else if (i.defaults && i.defaults.bin === ju) m.vbaraw = Wr(e, "xl/vbaProject.bin", true)
		}
		m.bookType = s ? "xlsb": "xlsx";
		return m
	}
	function lk(e, r) {
		var t = r || {};
		var a = "Workbook",
		n = Qe.find(e, a);
		try {
			a = "/!DataSpaces/Version";
			n = Qe.find(e, a);
			if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
			wo(n.content);
			a = "/!DataSpaces/DataSpaceMap";
			n = Qe.find(e, a);
			if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
			var i = To(n.content);
			if (i.length !== 1 || i[0].comps.length !== 1 || i[0].comps[0].t !== 0 || i[0].name !== "StrongEncryptionDataSpace" || i[0].comps[0].v !== "EncryptedPackage") throw new Error("ECMA-376 Encrypted file bad " + a);
			a = "/!DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace";
			n = Qe.find(e, a);
			if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
			var s = yo(n.content);
			if (s.length != 1 || s[0] != "StrongEncryptionTransform") throw new Error("ECMA-376 Encrypted file bad " + a);
			a = "/!DataSpaces/TransformInfo/StrongEncryptionTransform/!Primary";
			n = Qe.find(e, a);
			if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
			_o(n.content)
		} catch(f) {}
		a = "/EncryptionInfo";
		n = Qe.find(e, a);
		if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
		var l = Ao(n.content);
		a = "/EncryptedPackage";
		n = Qe.find(e, a);
		if (!n || !n.content) throw new Error("ECMA-376 Encrypted file missing " + a);
		if (l[0] == 4 && typeof decrypt_agile !== "undefined") return decrypt_agile(l[1], n.content, t.password || "", t);
		if (l[0] == 2 && typeof decrypt_std76 !== "undefined") return decrypt_std76(l[1], n.content, t.password || "", t);
		throw new Error("File is password-protected")
	}
	function ok(e, r) {
		if (e && !e.SSF) {
			e.SSF = Tr(q)
		}
		if (e && e.SSF) {
			$e();
			Ve(e.SSF);
			r.revssf = lr(e.SSF);
			r.revssf[e.SSF[65535]] = 0;
			r.ssf = e.SSF
		}
		r.rels = {};
		r.wbrels = {};
		r.Strings = [];
		r.Strings.Count = 0;
		r.Strings.Unique = 0;
		if (gv) r.revStrings = new Map;
		else {
			r.revStrings = {};
			r.revStrings.foo = [];
			delete r.revStrings.foo
		}
		var t = "bin";
		var a = true;
		var n = fi();
		rk(r = r || {});
		var i = Xr();
		var s = "",
		f = 0;
		r.cellXfs = [];
		yv(r.cellXfs, {},
		{
			revssf: {
				General: 0
			}
		});
		if (!e.Props) e.Props = {};
		s = "docProps/core.xml";
		$r(i, s, xi(e.Props, r));
		n.coreprops.push(s);
		vi(r.rels, 2, s, ci.CORE_PROPS);
		s = "docProps/app.xml";
		if (e.Props && e.Props.SheetNames) {} else if (!e.Workbook || !e.Workbook.Sheets) e.Props.SheetNames = e.SheetNames;
		else {
			var l = [];
			for (var o = 0; o < e.SheetNames.length; ++o) if ((e.Workbook.Sheets[o] || {}).Hidden != 2) l.push(e.SheetNames[o]);
			e.Props.SheetNames = l
		}
		e.Props.Worksheets = e.Props.SheetNames.length;
		$r(i, s, Ii(e.Props, r));
		n.extprops.push(s);
		vi(r.rels, 3, s, ci.EXT_PROPS);
		if (e.Custprops !== e.Props && ir(e.Custprops || {}).length > 0) {
			s = "docProps/custom.xml";
			$r(i, s, Di(e.Custprops, r));
			n.custprops.push(s);
			vi(r.rels, 4, s, ci.CUST_PROPS)
		}
		var c = ["SheetJ5"];
		r.tcid = 0;
		for (f = 1; f <= e.SheetNames.length; ++f) {
			var u = {
				"!id": {}
			};
			var h = e.Sheets[e.SheetNames[f - 1]];
			var d = (h || {})["!type"] || "sheet";
			switch (d) {
			case "chart":
				;
			default:
				s = "xl/worksheets/sheet" + f + "." + t;
				$r(i, s, Tm(f - 1, r, e, u));
				n.sheets.push(s);
				vi(r.wbrels, -1, "worksheets/sheet" + f + "." + t, ci.WS[0]);
			}
			if (h) {
				var v = h["!comments"];
				var p = false;
				var m = "";
				if (v && v.length > 0) {
					var b = false;
					v.forEach(function(e) {
						e[1].forEach(function(e) {
							if (e.T == true) b = true
						})
					});
					if (b) {
						m = "xl/threadedComments/threadedComment" + f + ".xml";
						$r(i, m, Uu(v, c, r));
						n.threadedcomments.push(m);
						vi(u, -1, "../threadedComments/threadedComment" + f + ".xml", ci.TCMNT)
					}
					m = "xl/comments" + f + "." + t;
					$r(i, m, Gu(v, r));
					n.comments.push(m);
					vi(u, -1, "../comments" + f + "." + t, ci.CMNT);
					p = true
				}
				if (h["!legacy"]) {
					if (p) $r(i, "xl/drawings/vmlDrawing" + f + ".vml", Nu(f, h["!comments"]))
				}
				delete h["!comments"];
				delete h["!legacy"]
			}
			if (u["!id"].rId1) $r(i, ui(s), di(u))
		}
		if (r.Strings != null && r.Strings.length > 0) {
			s = "xl/sharedStrings." + t;
			$r(i, s, mo(r.Strings, r));
			n.strs.push(s);
			vi(r.wbrels, -1, "sharedStrings." + t, ci.SST)
		}
		s = "xl/workbook." + t;
		$r(i, s, ab(e, r));
		n.workbooks.push(s);
		vi(r.rels, 1, s, ci.WB);
		s = "xl/theme/theme1.xml";
		var g = nu(e.Themes, r);
		$r(i, s, g);
		n.themes.push(s);
		vi(r.wbrels, -1, "theme/theme1.xml", ci.THEME);
		s = "xl/styles." + t;
		$r(i, s, jc(e, r));
		n.styles.push(s);
		vi(r.wbrels, -1, "styles." + t, ci.STY);
		if (e.vbaraw && a) {
			s = "xl/vbaProject.bin";
			$r(i, s, e.vbaraw);
			n.vba.push(s);
			vi(r.wbrels, -1, "vbaProject.bin", ci.VBA)
		}
		s = "xl/metadata." + t;
		$r(i, s, Tu());
		n.metadata.push(s);
		vi(r.wbrels, -1, "metadata." + t, ci.XLMETA);
		if (c.length > 1) {
			s = "xl/persons/person.xml";
			$r(i, s, Wu(c, r));
			n.people.push(s);
			vi(r.wbrels, -1, "persons/person.xml", ci.PEOPLE)
		}
		$r(i, "[Content_Types].xml", oi(n, r));
		$r(i, "_rels/.rels", di(r.rels));
		$r(i, "xl/_rels/workbook." + t + ".rels", di(r.wbrels));
		delete r.revssf;
		delete r.ssf;
		return i
	}
	function ck(e, r) {
		if (e && !e.SSF) {
			e.SSF = Tr(q)
		}
		if (e && e.SSF) {
			$e();
			Ve(e.SSF);
			r.revssf = lr(e.SSF);
			r.revssf[e.SSF[65535]] = 0;
			r.ssf = e.SSF
		}
		r.rels = {};
		r.wbrels = {};
		r.Strings = [];
		r.Strings.Count = 0;
		r.Strings.Unique = 0;
		if (gv) r.revStrings = new Map;
		else {
			r.revStrings = {};
			r.revStrings.foo = [];
			delete r.revStrings.foo
		}
		var t = "xml";
		var a = Zu.indexOf(r.bookType) > -1;
		var n = fi();
		rk(r = r || {});
		var i = Xr();
		var s = "",
		f = 0;
		r.cellXfs = [];
		yv(r.cellXfs, {},
		{
			revssf: {
				General: 0
			}
		});
		if (!e.Props) e.Props = {};
		s = "docProps/core.xml";
		$r(i, s, xi(e.Props, r));
		n.coreprops.push(s);
		vi(r.rels, 2, s, ci.CORE_PROPS);
		s = "docProps/app.xml";
		if (e.Props && e.Props.SheetNames) {} else if (!e.Workbook || !e.Workbook.Sheets) e.Props.SheetNames = e.SheetNames;
		else {
			var l = [];
			for (var o = 0; o < e.SheetNames.length; ++o) if ((e.Workbook.Sheets[o] || {}).Hidden != 2) l.push(e.SheetNames[o]);
			e.Props.SheetNames = l
		}
		e.Props.Worksheets = e.Props.SheetNames.length;
		$r(i, s, Ii(e.Props, r));
		n.extprops.push(s);
		vi(r.rels, 3, s, ci.EXT_PROPS);
		if (e.Custprops !== e.Props && ir(e.Custprops || {}).length > 0) {
			s = "docProps/custom.xml";
			$r(i, s, Di(e.Custprops, r));
			n.custprops.push(s);
			vi(r.rels, 4, s, ci.CUST_PROPS)
		}
		var c = ["SheetJ5"];
		r.tcid = 0;
		for (f = 1; f <= e.SheetNames.length; ++f) {
			var u = {
				"!id": {}
			};
			var h = e.Sheets[e.SheetNames[f - 1]];
			var d = (h || {})["!type"] || "sheet";
			switch (d) {
			case "chart":
				;
			default:
				s = "xl/worksheets/sheet" + f + "." + t;
				$r(i, s, ap(f - 1, r, e, u));
				n.sheets.push(s);
				vi(r.wbrels, -1, "worksheets/sheet" + f + "." + t, ci.WS[0]);
			}
			if (h) {
				var v = h["!comments"];
				var p = false;
				var m = "";
				if (v && v.length > 0) {
					var b = false;
					v.forEach(function(e) {
						e[1].forEach(function(e) {
							if (e.T == true) b = true
						})
					});
					if (b) {
						m = "xl/threadedComments/threadedComment" + f + ".xml";
						$r(i, m, Uu(v, c, r));
						n.threadedcomments.push(m);
						vi(u, -1, "../threadedComments/threadedComment" + f + ".xml", ci.TCMNT)
					}
					m = "xl/comments" + f + "." + t;
					$r(i, m, Lu(v, r));
					n.comments.push(m);
					vi(u, -1, "../comments" + f + "." + t, ci.CMNT);
					p = true
				}
				if (h["!legacy"]) {
					if (p) $r(i, "xl/drawings/vmlDrawing" + f + ".vml", Nu(f, h["!comments"]))
				}
				delete h["!comments"];
				delete h["!legacy"]
			}
			if (u["!id"].rId1) $r(i, ui(s), di(u))
		}
		if (r.Strings != null && r.Strings.length > 0) {
			s = "xl/sharedStrings." + t;
			$r(i, s, co(r.Strings, r));
			n.strs.push(s);
			vi(r.wbrels, -1, "sharedStrings." + t, ci.SST)
		}
		s = "xl/workbook." + t;
		$r(i, s, zm(e, r));
		n.workbooks.push(s);
		vi(r.rels, 1, s, ci.WB);
		s = "xl/theme/theme1.xml";
		$r(i, s, nu(e.Themes, r));
		n.themes.push(s);
		vi(r.wbrels, -1, "theme/theme1.xml", ci.THEME);
		s = "xl/styles." + t;
		$r(i, s, kc(e, r));
		n.styles.push(s);
		vi(r.wbrels, -1, "styles." + t, ci.STY);
		if (e.vbaraw && a) {
			s = "xl/vbaProject.bin";
			$r(i, s, e.vbaraw);
			n.vba.push(s);
			vi(r.wbrels, -1, "vbaProject.bin", ci.VBA)
		}
		s = "xl/metadata." + t;
		$r(i, s, Eu());
		n.metadata.push(s);
		vi(r.wbrels, -1, "metadata." + t, ci.XLMETA);
		if (c.length > 1) {
			s = "xl/persons/person.xml";
			$r(i, s, Wu(c, r));
			n.people.push(s);
			vi(r.wbrels, -1, "persons/person.xml", ci.PEOPLE)
		}
		$r(i, "[Content_Types].xml", oi(n, r));
		$r(i, "_rels/.rels", di(r.rels));
		$r(i, "xl/_rels/workbook." + t + ".rels", di(r.wbrels));
		delete r.revssf;
		delete r.ssf;
		return i
	}
	function uk(e, r) {
		var t = "";
		switch ((r || {}).type || "base64") {
		case "buffer":
			return [e[0], e[1], e[2], e[3], e[4], e[5], e[6], e[7]];
		case "base64":
			t = _(e.slice(0, 12));
			break;
		case "binary":
			t = e;
			break;
		case "array":
			return [e[0], e[1], e[2], e[3], e[4], e[5], e[6], e[7]];
		default:
			throw new Error("Unrecognized type " + (r && r.type || "undefined"));
		}
		return [t.charCodeAt(0), t.charCodeAt(1), t.charCodeAt(2), t.charCodeAt(3), t.charCodeAt(4), t.charCodeAt(5), t.charCodeAt(6), t.charCodeAt(7)]
	}
	function hk(e, r) {
		if (Qe.find(e, "EncryptedPackage")) return lk(e, r);
		return Jb(e, r)
	}
	function dk(e, r) {
		var t, a = e;
		var n = r || {};
		if (!n.type) n.type = S && Buffer.isBuffer(e) ? "buffer": "base64";
		t = Gr(a, n);
		return fk(t, n)
	}
	function vk(e, r) {
		var t = 0;
		e: while (t < e.length) switch (e.charCodeAt(t)) {
		case 10:
			;
		case 13:
			;
		case 32:
			++t;
			break;
		case 60:
			return Cb(e.slice(t), r);
		default:
			break e;
		}
		return Yl.to_workbook(e, r)
	}
	function pk(e, r) {
		var t = "",
		a = uk(e, r);
		switch (r.type) {
		case "base64":
			t = _(e);
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
		return vk(t, r)
	}
	function mk(e, r) {
		var t = e;
		if (r.type == "base64") t = _(t);
		if (typeof ArrayBuffer !== "undefined" && e instanceof ArrayBuffer) t = new Uint8Array(e);
		t = typeof a !== "undefined" ? a.utils.decode(1200, t.slice(2), "str") : S && Buffer.isBuffer(e) ? e.slice(2).toString("utf16le") : typeof Uint8Array !== "undefined" && t instanceof Uint8Array ? typeof TextDecoder !== "undefined" ? new TextDecoder("utf-16le").decode(t.slice(2)) : h(t.slice(2)) : u(t.slice(2));
		r.type = "binary";
		return vk(t, r)
	}
	function bk(e) {
		return ! e.match(/[^\x00-\x7F]/) ? e: Tt(e)
	}
	function gk(e, r, t, a) {
		if (a) {
			t.type = "string";
			return Yl.to_workbook(e, t)
		}
		return Yl.to_workbook(r, t)
	}
	function wk(e, r) {
		o();
		var t = r || {};
		if (t.codepage && typeof a === "undefined") console.error("Codepage tables are not loaded.  Non-ASCII characters may not give expected results");
		if (typeof ArrayBuffer !== "undefined" && e instanceof ArrayBuffer) return wk(new Uint8Array(e), (t = Tr(t), t.type = "array", t));
		if (typeof Uint8Array !== "undefined" && e instanceof Uint8Array && !t.type) t.type = typeof Deno !== "undefined" ? "buffer": "array";
		var n = e,
		i = [0, 0, 0, 0],
		s = false;
		if (t.cellStyles) {
			t.cellNF = true;
			t.sheetStubs = true
		}
		bv = {};
		if (t.dateNF) bv.dateNF = t.dateNF;
		if (!t.type) t.type = S && Buffer.isBuffer(e) ? "buffer": "base64";
		if (t.type == "file") {
			t.type = S ? "buffer": "binary";
			n = nr(e);
			if (typeof Uint8Array !== "undefined" && !S) t.type = "array"
		}
		if (t.type == "string") {
			s = true;
			t.type = "binary";
			t.codepage = 65001;
			n = bk(e)
		}
		if (t.type == "array" && typeof Uint8Array !== "undefined" && e instanceof Uint8Array && typeof ArrayBuffer !== "undefined") {
			var f = new ArrayBuffer(3),
			l = new Uint8Array(f);
			l.foo = "bar";
			if (!l.foo) {
				t = Tr(t);
				t.type = "array";
				return wk(D(n), t)
			}
		}
		switch ((i = uk(n, t))[0]) {
		case 208:
			if (i[1] === 207 && i[2] === 17 && i[3] === 224 && i[4] === 161 && i[5] === 177 && i[6] === 26 && i[7] === 225) return hk(Qe.read(n, t), t);
			break;
		case 9:
			if (i[1] <= 8) return Jb(n, t);
			break;
		case 60:
			return Cb(n, t);
		case 73:
			if (i[1] === 73 && i[2] === 42 && i[3] === 0) throw new Error("TIFF Image File is not a spreadsheet");
			if (i[1] === 68) return Zl(n, t);
			break;
		case 84:
			if (i[1] === 65 && i[2] === 66 && i[3] === 76) return jl.to_workbook(n, t);
			break;
		case 80:
			return i[1] === 75 && i[2] < 9 && i[3] < 9 ? dk(n, t) : gk(e, n, t, s);
		case 239:
			return i[3] === 60 ? Cb(n, t) : gk(e, n, t, s);
		case 255:
			if (i[1] === 254) {
				return mk(n, t)
			} else if (i[1] === 0 && i[2] === 2 && i[3] === 0) return Jl.to_workbook(n, t);
			break;
		case 0:
			if (i[1] === 0) {
				if (i[2] >= 2 && i[3] === 0) return Jl.to_workbook(n, t);
				if (i[2] === 0 && (i[3] === 8 || i[3] === 9)) return Jl.to_workbook(n, t)
			}
			break;
		case 3:
			;
		case 131:
			;
		case 139:
			;
		case 140:
			return Xl.to_workbook(n, t);
		case 123:
			if (i[1] === 92 && i[2] === 114 && i[3] === 116) return Ho(n, t);
			break;
		case 10:
			;
		case 13:
			;
		case 32:
			return pk(n, t);
		case 137:
			if (i[1] === 80 && i[2] === 78 && i[3] === 71) throw new Error("PNG Image File is not a spreadsheet");
			break;
		case 8:
			if (i[1] === 231) throw new Error("Unsupported Multiplan 1.x file!");
			break;
		case 12:
			if (i[1] === 236) throw new Error("Unsupported Multiplan 2.x file!");
			if (i[1] === 237) throw new Error("Unsupported Multiplan 3.x file!");
			break;
		}
		if ($l.indexOf(i[0]) > -1 && i[2] <= 12 && i[3] <= 31) return Xl.to_workbook(n, t);
		return gk(e, n, t, s)
	}
	function kk(e, r) {
		var t = r || {};
		t.type = "file";
		return wk(e, t)
	}
	function Tk(e, r) {
		switch (r.type) {
		case "base64":
			;
		case "binary":
			break;
		case "buffer":
			;
		case "array":
			r.type = "";
			break;
		case "file":
			return ar(r.file, Qe.write(e, {
				type: S ? "buffer": ""
			}));
		case "string":
			throw new Error("'string' output type invalid for '" + r.bookType + "' files");
		default:
			throw new Error("Unrecognized type " + r.type);
		}
		return Qe.write(e, r)
	}
	function yk(e, r) {
		switch (r.bookType) {
		case "ods":
			return Xg(e, r);
		case "numbers":
			return Hw(e, r);
		case "xlsb":
			return ok(e, r);
		default:
			return ck(e, r);
		}
	}
	function Ek(e, r) {
		var t = Tr(r || {});
		var a = yk(e, t);
		return Sk(a, t)
	}
	function _k(e, r) {
		var t = Tr(r || {});
		var a = ck(e, t);
		return Sk(a, t)
	}
	function Sk(e, r) {
		var t = {};
		var a = S ? "nodebuffer": typeof Uint8Array !== "undefined" ? "array": "string";
		if (r.compression) t.compression = "DEFLATE";
		if (r.password) t.type = a;
		else switch (r.type) {
		case "base64":
			t.type = "base64";
			break;
		case "binary":
			t.type = "string";
			break;
		case "string":
			throw new Error("'string' output type invalid for '" + r.bookType + "' files");
		case "buffer":
			;
		case "file":
			t.type = a;
			break;
		default:
			throw new Error("Unrecognized type " + r.type);
		}
		var n = e.FullPaths ? Qe.write(e, {
			fileType: "zip",
			type: {
				nodebuffer: "buffer",
				string: "binary"
			} [t.type] || t.type,
			compression: !!r.compression
		}) : e.generate(t);
		if (typeof Deno !== "undefined") {
			if (typeof n == "string") {
				if (r.type == "binary" || r.type == "base64") return n;
				n = new Uint8Array(I(n))
			}
		}
		if (r.password && typeof encrypt_agile !== "undefined") return Tk(encrypt_agile(n, r.password), r);
		if (r.type === "file") return ar(r.file, n);
		return r.type == "string" ? kt(n) : n
	}
	function xk(e, r) {
		var t = r || {};
		var a = qb(e, t);
		return Tk(a, t)
	}
	function Ak(e, r, t) {
		if (!t) t = "";
		var a = t + e;
		switch (r.type) {
		case "base64":
			return T(Tt(a));
		case "binary":
			return Tt(a);
		case "string":
			return e;
		case "file":
			return ar(r.file, a, "utf8");
		case "buffer":
			{
				if (S) return x(a, "utf8");
				else if (typeof TextEncoder !== "undefined") return (new TextEncoder).encode(a);
				else return Ak(a, {
					type: "binary"
				}).split("").map(function(e) {
					return e.charCodeAt(0)
				})
			};
		}
		throw new Error("Unrecognized type " + r.type)
	}
	function Ck(e, r) {
		switch (r.type) {
		case "base64":
			return y(e);
		case "binary":
			return e;
		case "string":
			return e;
		case "file":
			return ar(r.file, e, "binary");
		case "buffer":
			{
				if (S) return x(e, "binary");
				else return e.split("").map(function(e) {
					return e.charCodeAt(0)
				})
			};
		}
		throw new Error("Unrecognized type " + r.type)
	}
	function Rk(e, r) {
		switch (r.type) {
		case "string":
			;
		case "base64":
			;
		case "binary":
			var t = "";
			for (var a = 0; a < e.length; ++a) t += String.fromCharCode(e[a]);
			return r.type == "base64" ? T(t) : r.type == "string" ? kt(t) : t;
		case "file":
			return ar(r.file, e);
		case "buffer":
			return e;
		default:
			throw new Error("Unrecognized type " + r.type);
		}
	}
	function Ok(e, r) {
		o();
		Um(e);
		var t = Tr(r || {});
		if (t.cellStyles) {
			t.cellNF = true;
			t.sheetStubs = true
		}
		if (t.type == "array") {
			t.type = "binary";
			var a = Ok(e, t);
			t.type = "array";
			return I(a)
		}
		return _k(e, t)
	}
	function Ik(e, r) {
		o();
		Um(e);
		var t = Tr(r || {});
		if (t.cellStyles) {
			t.cellNF = true;
			t.sheetStubs = true
		}
		if (t.type == "array") {
			t.type = "binary";
			var a = Ik(e, t);
			t.type = "array";
			return I(a)
		}
		var n = 0;
		if (t.sheet) {
			if (typeof t.sheet == "number") n = t.sheet;
			else n = e.SheetNames.indexOf(t.sheet);
			if (!e.SheetNames[n]) throw new Error("Sheet not found: " + t.sheet + " : " + typeof t.sheet)
		}
		switch (t.bookType || "xlsb") {
		case "xml":
			;
		case "xlml":
			return Ak(zb(e, t), t);
		case "slk":
			;
		case "sylk":
			return Ak(Gl.from_sheet(e.Sheets[e.SheetNames[n]], t, e), t);
		case "htm":
			;
		case "html":
			return Ak(Og(e.Sheets[e.SheetNames[n]], t), t);
		case "txt":
			return Ck(zk(e.Sheets[e.SheetNames[n]], t), t);
		case "csv":
			return Ak(Wk(e.Sheets[e.SheetNames[n]], t), t, "\ufeff");
		case "dif":
			return Ak(jl.from_sheet(e.Sheets[e.SheetNames[n]], t), t);
		case "dbf":
			return Rk(Xl.from_sheet(e.Sheets[e.SheetNames[n]], t), t);
		case "prn":
			return Ak(Yl.from_sheet(e.Sheets[e.SheetNames[n]], t), t);
		case "rtf":
			return Ak(Vo(e.Sheets[e.SheetNames[n]], t), t);
		case "eth":
			return Ak(Kl.from_sheet(e.Sheets[e.SheetNames[n]], t), t);
		case "fods":
			return Ak(Xg(e, t), t);
		case "wk1":
			return Rk(Jl.sheet_to_wk1(e.Sheets[e.SheetNames[n]], t), t);
		case "wk3":
			return Rk(Jl.book_to_wk3(e, t), t);
		case "biff2":
			if (!t.biff) t.biff = 2;
		case "biff3":
			if (!t.biff) t.biff = 3;
		case "biff4":
			if (!t.biff) t.biff = 4;
			return Rk(Eg(e, t), t);
		case "biff5":
			if (!t.biff) t.biff = 5;
		case "biff8":
			;
		case "xla":
			;
		case "xls":
			if (!t.biff) t.biff = 8;
			return xk(e, t);
		case "xlsx":
			;
		case "xlsm":
			;
		case "xlam":
			;
		case "xlsb":
			;
		case "numbers":
			;
		case "ods":
			return Ek(e, t);
		default:
			throw new Error("Unrecognized bookType |" + t.bookType + "|");
		}
	}
	function Nk(e) {
		if (e.bookType) return;
		var r = {
			xls: "biff8",
			htm: "html",
			slk: "sylk",
			socialcalc: "eth",
			Sh33tJS: "WTF"
		};
		var t = e.file.slice(e.file.lastIndexOf(".")).toLowerCase();
		if (t.match(/^\.[a-z]+$/)) e.bookType = t.slice(1);
		e.bookType = r[e.bookType] || e.bookType
	}
	function Fk(e, r, t) {
		var a = t || {};
		a.type = "file";
		a.file = r;
		Nk(a);
		return Ik(e, a)
	}
	function Dk(e, r, t) {
		var a = t || {};
		a.type = "file";
		a.file = r;
		Nk(a);
		return Ok(e, a)
	}
	function Pk(e, r, t, a) {
		var n = t || {};
		n.type = "file";
		n.file = e;
		Nk(n);
		n.type = "buffer";
		var i = a;
		if (! (i instanceof Function)) i = t;
		return er.writeFile(e, Ik(r, n), i)
	}
	function Lk(e, r, t, a, n, i, s) {
		var f = Na(t);
		var l = s.defval,
		o = s.raw || !Object.prototype.hasOwnProperty.call(s, "raw");
		var c = true,
		u = e["!data"] != null;
		var h = n === 1 ? [] : {};
		if (n !== 1) {
			if (Object.defineProperty) try {
				Object.defineProperty(h, "__rowNum__", {
					value: t,
					enumerable: false
				})
			} catch(d) {
				h.__rowNum__ = t
			} else h.__rowNum__ = t
		}
		if (!u || e["!data"][t]) for (var v = r.s.c; v <= r.e.c; ++v) {
			var p = u ? (e["!data"][t] || [])[v] : e[a[v] + f];
			if (p == null || p.t === undefined) {
				if (l === undefined) continue;
				if (i[v] != null) {
					h[i[v]] = l
				}
				continue
			}
			var m = p.v;
			switch (p.t) {
			case "z":
				if (m == null) break;
				continue;
			case "e":
				m = m == 0 ? null: void 0;
				break;
			case "s":
				;
			case "b":
				;
			case "n":
				if (!p.z || !Le(p.z)) break;
				m = vr(m);
				if (typeof m == "number") break;
			case "d":
				if (! (s && (s.UTC || s.raw === false))) m = Fr(new Date(m));
				break;
			default:
				throw new Error("unrecognized type " + p.t);
			}
			if (i[v] != null) {
				if (m == null) {
					if (p.t == "e" && m === null) h[i[v]] = null;
					else if (l !== undefined) h[i[v]] = l;
					else if (o && m === null) h[i[v]] = null;
					else continue
				} else {
					h[i[v]] = (p.t === "n" && typeof s.rawNumbers === "boolean" ? s.rawNumbers: o) ? m: Ka(p, m, s)
				}
				if (m != null) c = false
			}
		}
		return {
			row: h,
			isempty: c
		}
	}
	function Mk(e, r) {
		if (e == null || e["!ref"] == null) return [];
		var t = {
			t: "n",
			v: 0
		},
		a = 0,
		n = 1,
		i = [],
		s = 0,
		f = "";
		var l = {
			s: {
				r: 0,
				c: 0
			},
			e: {
				r: 0,
				c: 0
			}
		};
		var o = r || {};
		var c = o.range != null ? o.range: e["!ref"];
		if (o.header === 1) a = 1;
		else if (o.header === "A") a = 2;
		else if (Array.isArray(o.header)) a = 3;
		else if (o.header == null) a = 0;
		switch (typeof c) {
		case "string":
			l = Ga(c);
			break;
		case "number":
			l = Ga(e["!ref"]);
			l.s.r = c;
			break;
		default:
			l = c;
		}
		if (a > 0) n = 0;
		var u = Na(l.s.r);
		var h = [];
		var d = [];
		var v = 0,
		p = 0;
		var m = e["!data"] != null;
		var b = l.s.r,
		g = 0;
		var w = {};
		if (m && !e["!data"][b]) e["!data"][b] = [];
		var k = o.skipHidden && e["!cols"] || [];
		var T = o.skipHidden && e["!rows"] || [];
		for (g = l.s.c; g <= l.e.c; ++g) {
			if ((k[g] || {}).hidden) continue;
			h[g] = La(g);
			t = m ? e["!data"][b][g] : e[h[g] + u];
			switch (a) {
			case 1:
				i[g] = g - l.s.c;
				break;
			case 2:
				i[g] = h[g];
				break;
			case 3:
				i[g] = o.header[g - l.s.c];
				break;
			default:
				if (t == null) t = {
					w: "__EMPTY",
					t: "s"
				};
				f = s = Ka(t, null, o);
				p = w[s] || 0;
				if (!p) w[s] = 1;
				else {
					do {
						f = s + "_" + p++
					} while ( w [ f ]);
					w[s] = p;
					w[f] = 1
				}
				i[g] = f;
			}
		}
		for (b = l.s.r + n; b <= l.e.r; ++b) {
			if ((T[b] || {}).hidden) continue;
			var y = Lk(e, l, b, h, a, i, o);
			if (y.isempty === false || (a === 1 ? o.blankrows !== false: !!o.blankrows)) d[v++] = y.row
		}
		d.length = v;
		return d
	}
	var Uk = /"/g;
	function Bk(e, r, t, a, n, i, s, f) {
		var l = true;
		var o = [],
		c = "",
		u = Na(t);
		var h = e["!data"] != null;
		var d = h && e["!data"][t] || [];
		for (var v = r.s.c; v <= r.e.c; ++v) {
			if (!a[v]) continue;
			var p = h ? d[v] : e[a[v] + u];
			if (p == null) c = "";
			else if (p.v != null) {
				l = false;
				c = "" + (f.rawNumbers && p.t == "n" ? p.v: Ka(p, null, f));
				for (var m = 0,
				b = 0; m !== c.length; ++m) if ((b = c.charCodeAt(m)) === n || b === i || b === 34 || f.forceQuotes) {
					c = '"' + c.replace(Uk, '""') + '"';
					break
				}
				if (c == "ID") c = '"ID"'
			} else if (p.f != null && !p.F) {
				l = false;
				c = "=" + p.f;
				if (c.indexOf(",") >= 0) c = '"' + c.replace(Uk, '""') + '"'
			} else c = "";
			o.push(c)
		}
		if (f.blankrows === false && l) return null;
		return o.join(s)
	}
	function Wk(e, r) {
		var t = [];
		var a = r == null ? {}: r;
		if (e == null || e["!ref"] == null) return "";
		var n = Ga(e["!ref"]);
		var i = a.FS !== undefined ? a.FS: ",",
		s = i.charCodeAt(0);
		var f = a.RS !== undefined ? a.RS: "\n",
		l = f.charCodeAt(0);
		var o = new RegExp((i == "|" ? "\\|": i) + "+$");
		var c = "",
		u = [];
		var h = a.skipHidden && e["!cols"] || [];
		var d = a.skipHidden && e["!rows"] || [];
		for (var v = n.s.c; v <= n.e.c; ++v) if (! (h[v] || {}).hidden) u[v] = La(v);
		var p = 0;
		for (var m = n.s.r; m <= n.e.r; ++m) {
			if ((d[m] || {}).hidden) continue;
			c = Bk(e, n, m, u, s, l, i, a);
			if (c == null) {
				continue
			}
			if (a.strip) c = c.replace(o, "");
			if (c || a.blankrows !== false) t.push((p++?f: "") + c)
		}
		return t.join("")
	}
	function zk(e, r) {
		if (!r) r = {};
		r.FS = "\t";
		r.RS = "\n";
		var t = Wk(e, r);
		if (typeof a == "undefined" || r.type == "string") return t;
		var n = a.utils.encode(1200, t, "str");
		return String.fromCharCode(255) + String.fromCharCode(254) + n
	}
	function Hk(e) {
		var r = "",
		t, a = "";
		if (e == null || e["!ref"] == null) return [];
		var n = Ga(e["!ref"]),
		i = "",
		s = [],
		f;
		var l = [];
		var o = e["!data"] != null;
		for (f = n.s.c; f <= n.e.c; ++f) s[f] = La(f);
		for (var c = n.s.r; c <= n.e.r; ++c) {
			i = Na(c);
			for (f = n.s.c; f <= n.e.c; ++f) {
				r = s[f] + i;
				t = o ? (e["!data"][c] || [])[f] : e[r];
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
				l[l.length] = r + "=" + a
			}
		}
		return l
	}
	function Vk(e, r, t) {
		var a = t || {};
		var n = e ? e["!data"] != null: a.dense;
		if (g != null && n == null) n = g;
		var i = +!a.skipHeader;
		var s = e || {};
		if (!e && n) s["!data"] = [];
		var f = 0,
		l = 0;
		if (s && a.origin != null) {
			if (typeof a.origin == "number") f = a.origin;
			else {
				var o = typeof a.origin == "string" ? Wa(a.origin) : a.origin;
				f = o.r;
				l = o.c
			}
		}
		var c = {
			s: {
				c: 0,
				r: 0
			},
			e: {
				c: l,
				r: f + r.length - 1 + i
			}
		};
		if (s["!ref"]) {
			var u = Ga(s["!ref"]);
			c.e.c = Math.max(c.e.c, u.e.c);
			c.e.r = Math.max(c.e.r, u.e.r);
			if (f == -1) {
				f = u.e.r + 1;
				c.e.r = f + r.length - 1 + i
			}
		} else {
			if (f == -1) {
				f = 0;
				c.e.r = r.length - 1 + i
			}
		}
		var h = a.header || [],
		d = 0;
		var v = [];
		r.forEach(function(e, r) {
			if (n && !s["!data"][f + r + i]) s["!data"][f + r + i] = [];
			if (n) v = s["!data"][f + r + i];
			ir(e).forEach(function(t) {
				if ((d = h.indexOf(t)) == -1) h[d = h.length] = t;
				var o = e[t];
				var c = "z";
				var u = "";
				var p = n ? "": La(l + d) + Na(f + r + i);
				var m = n ? v[l + d] : s[p];
				if (o && typeof o === "object" && !(o instanceof Date)) {
					if (n) v[l + d] = o;
					else s[p] = o
				} else {
					if (typeof o == "number") c = "n";
					else if (typeof o == "boolean") c = "b";
					else if (typeof o == "string") c = "s";
					else if (o instanceof Date) {
						c = "d";
						if (!a.UTC) o = Dr(o);
						if (!a.cellDates) {
							c = "n";
							o = dr(o)
						}
						u = m != null && m.z && Le(m.z) ? m.z: a.dateNF || q[14]
					} else if (o === null && a.nullError) {
						c = "e";
						o = 0
					}
					if (!m) {
						if (!n) s[p] = m = {
							t: c,
							v: o
						};
						else v[l + d] = m = {
							t: c,
							v: o
						}
					} else {
						m.t = c;
						m.v = o;
						delete m.w;
						delete m.R;
						if (u) m.z = u
					}
					if (u) m.z = u
				}
			})
		});
		c.e.c = Math.max(c.e.c, l + h.length - 1);
		var p = Na(f);
		if (n && !s["!data"][f]) s["!data"][f] = [];
		if (i) for (d = 0; d < h.length; ++d) {
			if (n) s["!data"][f][d + l] = {
				t: "s",
				v: h[d]
			};
			else s[La(d + l) + p] = {
				t: "s",
				v: h[d]
			}
		}
		s["!ref"] = Va(c);
		return s
	}
	function $k(e, r) {
		return Vk(null, e, r)
	}
	function Xk(e, r, t) {
		if (typeof r == "string") {
			if (e["!data"] != null) {
				var a = Wa(r);
				if (!e["!data"][a.r]) e["!data"][a.r] = [];
				return e["!data"][a.r][a.c] || (e["!data"][a.r][a.c] = {
					t: "z"
				})
			}
			return e[r] || (e[r] = {
				t: "z"
			})
		}
		if (typeof r != "number") return Xk(e, za(r));
		return Xk(e, La(t || 0) + Na(r))
	}
	function Gk(e, r) {
		if (typeof r == "number") {
			if (r >= 0 && e.SheetNames.length > r) return r;
			throw new Error("Cannot find sheet # " + r)
		} else if (typeof r == "string") {
			var t = e.SheetNames.indexOf(r);
			if (t > -1) return t;
			throw new Error("Cannot find sheet name |" + r + "|")
		} else throw new Error("Cannot find sheet |" + r + "|")
	}
	function jk(e, r) {
		var t = {
			SheetNames: [],
			Sheets: {}
		};
		if (e) Kk(t, e, r || "Sheet1");
		return t
	}
	function Kk(e, r, t, a) {
		var n = 1;
		if (!t) for (; n <= 65535; ++n, t = undefined) if (e.SheetNames.indexOf(t = "Sheet" + n) == -1) break;
		if (!t || e.SheetNames.length >= 65535) throw new Error("Too many worksheets");
		if (a && e.SheetNames.indexOf(t) >= 0) {
			var i = t.match(/(^.*?)(\d+)$/);
			n = i && +i[2] || 0;
			var s = i && i[1] || t;
			for (++n; n <= 65535; ++n) if (e.SheetNames.indexOf(t = s + n) == -1) break
		}
		Lm(t);
		if (e.SheetNames.indexOf(t) >= 0) throw new Error("Worksheet with name |" + t + "| already exists!");
		e.SheetNames.push(t);
		e.Sheets[t] = r;
		return t
	}
	function Yk(e, r, t) {
		if (!e.Workbook) e.Workbook = {};
		if (!e.Workbook.Sheets) e.Workbook.Sheets = [];
		var a = Gk(e, r);
		if (!e.Workbook.Sheets[a]) e.Workbook.Sheets[a] = {};
		switch (t) {
		case 0:
			;
		case 1:
			;
		case 2:
			break;
		default:
			throw new Error("Bad sheet visibility setting " + t);
		}
		e.Workbook.Sheets[a].Hidden = t
	}
	function Zk(e, r) {
		e.z = r;
		return e
	}
	function Jk(e, r, t) {
		if (!r) {
			delete e.l
		} else {
			e.l = {
				Target: r
			};
			if (t) e.l.Tooltip = t
		}
		return e
	}
	function qk(e, r, t) {
		return Jk(e, "#" + r, t)
	}
	function Qk(e, r, t) {
		if (!e.c) e.c = [];
		e.c.push({
			t: r,
			a: t || "SheetJS"
		})
	}
	function eT(e, r, t, a) {
		var n = typeof r != "string" ? r: Ga(r);
		var i = typeof r == "string" ? r: Va(r);
		for (var s = n.s.r; s <= n.e.r; ++s) for (var f = n.s.c; f <= n.e.c; ++f) {
			var l = Xk(e, s, f);
			l.t = "n";
			l.F = i;
			delete l.v;
			if (s == n.s.r && f == n.s.c) {
				l.f = t;
				if (a) l.D = true
			}
		}
		var o = Ha(e["!ref"]);
		if (o.s.r > n.s.r) o.s.r = n.s.r;
		if (o.s.c > n.s.c) o.s.c = n.s.c;
		if (o.e.r < n.e.r) o.e.r = n.e.r;
		if (o.e.c < n.e.c) o.e.c = n.e.c;
		e["!ref"] = Va(o);
		return e
	}
	var rT = {
		encode_col: La,
		encode_row: Na,
		encode_cell: za,
		encode_range: Va,
		decode_col: Pa,
		decode_row: Ia,
		split_cell: Ba,
		decode_cell: Wa,
		decode_range: Ha,
		format_cell: Ka,
		sheet_new: Za,
		sheet_add_aoa: Ja,
		sheet_add_json: Vk,
		sheet_add_dom: Ig,
		aoa_to_sheet: qa,
		json_to_sheet: $k,
		table_to_sheet: Ng,
		table_to_book: Fg,
		sheet_to_csv: Wk,
		sheet_to_txt: zk,
		sheet_to_json: Mk,
		sheet_to_html: Og,
		sheet_to_formulae: Hk,
		sheet_to_row_object_array: Mk,
		sheet_get_cell: Xk,
		book_new: jk,
		book_append_sheet: Kk,
		book_set_sheet_visibility: Yk,
		cell_set_number_format: Zk,
		cell_set_hyperlink: Jk,
		cell_set_internal_link: qk,
		cell_add_comment: Qk,
		sheet_set_array_formula: eT,
		consts: {
			SHEET_VISIBLE: 0,
			SHEET_HIDDEN: 1,
			SHEET_VERY_HIDDEN: 2
		}
	};
	var tT;
	function aT(e) {
		tT = e
	}
	function nT(e, r) {
		var t = tT();
		var a = r == null ? {}: r;
		if (e == null || e["!ref"] == null) {
			t.push(null);
			return t
		}
		var n = Ga(e["!ref"]);
		var i = a.FS !== undefined ? a.FS: ",",
		s = i.charCodeAt(0);
		var f = a.RS !== undefined ? a.RS: "\n",
		l = f.charCodeAt(0);
		var o = new RegExp((i == "|" ? "\\|": i) + "+$");
		var c = "",
		u = [];
		var h = a.skipHidden && e["!cols"] || [];
		var d = a.skipHidden && e["!rows"] || [];
		for (var v = n.s.c; v <= n.e.c; ++v) if (! (h[v] || {}).hidden) u[v] = La(v);
		var p = n.s.r;
		var m = false,
		b = 0;
		t._read = function() {
			if (!m) {
				m = true;
				return t.push("\ufeff")
			}
			while (p <= n.e.r) {++p;
				if ((d[p - 1] || {}).hidden) continue;
				c = Bk(e, n, p - 1, u, s, l, i, a);
				if (c != null) {
					if (a.strip) c = c.replace(o, "");
					if (c || a.blankrows !== false) return t.push((b++?f: "") + c)
				}
			}
			return t.push(null)
		};
		return t
	}
	function iT(e, r) {
		var t = tT();
		var a = r || {};
		var n = a.header != null ? a.header: xg;
		var i = a.footer != null ? a.footer: Ag;
		t.push(n);
		var s = Ha(e["!ref"]);
		t.push(Rg(e, s, a));
		var f = s.s.r;
		var l = false;
		t._read = function() {
			if (f > s.e.r) {
				if (!l) {
					l = true;
					t.push("</table>" + i)
				}
				return t.push(null)
			}
			while (f <= s.e.r) {
				t.push(Sg(e, s, f, a)); ++f;
				break
			}
		};
		return t
	}
	function sT(e, r) {
		var t = tT({
			objectMode: true
		});
		if (e == null || e["!ref"] == null) {
			t.push(null);
			return t
		}
		var a = {
			t: "n",
			v: 0
		},
		n = 0,
		i = 1,
		s = [],
		f = 0,
		l = "";
		var o = {
			s: {
				r: 0,
				c: 0
			},
			e: {
				r: 0,
				c: 0
			}
		};
		var c = r || {};
		var u = c.range != null ? c.range: e["!ref"];
		if (c.header === 1) n = 1;
		else if (c.header === "A") n = 2;
		else if (Array.isArray(c.header)) n = 3;
		switch (typeof u) {
		case "string":
			o = Ga(u);
			break;
		case "number":
			o = Ga(e["!ref"]);
			o.s.r = u;
			break;
		default:
			o = u;
		}
		if (n > 0) i = 0;
		var h = Na(o.s.r);
		var d = [];
		var v = 0;
		var p = e["!data"] != null;
		var m = o.s.r,
		b = 0;
		var g = {};
		if (p && !e["!data"][m]) e["!data"][m] = [];
		var w = c.skipHidden && e["!cols"] || [];
		var k = c.skipHidden && e["!rows"] || [];
		for (b = o.s.c; b <= o.e.c; ++b) {
			if ((w[b] || {}).hidden) continue;
			d[b] = La(b);
			a = p ? e["!data"][m][b] : e[d[b] + h];
			switch (n) {
			case 1:
				s[b] = b - o.s.c;
				break;
			case 2:
				s[b] = d[b];
				break;
			case 3:
				s[b] = c.header[b - o.s.c];
				break;
			default:
				if (a == null) a = {
					w: "__EMPTY",
					t: "s"
				};
				l = f = Ka(a, null, c);
				v = g[f] || 0;
				if (!v) g[f] = 1;
				else {
					do {
						l = f + "_" + v++
					} while ( g [ l ]);
					g[f] = v;
					g[l] = 1
				}
				s[b] = l;
			}
		}
		m = o.s.r + i;
		t._read = function() {
			while (m <= o.e.r) {
				if ((k[m - 1] || {}).hidden) continue;
				var r = Lk(e, o, m, d, n, s, c); ++m;
				if (r.isempty === false || (n === 1 ? c.blankrows !== false: !!c.blankrows)) {
					t.push(r.row);
					return
				}
			}
			return t.push(null)
		};
		return t
	}
	var fT = {
		to_json: sT,
		to_html: iT,
		to_csv: nT,
		set_readable: aT
	};
	if (typeof Jb !== "undefined") e.parse_xlscfb = Jb;
	e.parse_zip = fk;
	e.read = wk;
	e.readFile = kk;
	e.readFileSync = kk;
	e.write = Ik;
	e.writeFile = Fk;
	e.writeFileSync = Fk;
	e.writeFileAsync = Pk;
	e.utils = rT;
	e.writeXLSX = Ok;
	e.writeFileXLSX = Dk;
	e.set_fs = rr;
	e.set_cptable = b;
	e.SSF = Xe;
	if (typeof fT !== "undefined") e.stream = fT;
	if (typeof Qe !== "undefined") e.CFB = Qe;
	if (typeof require !== "undefined") {
		var lT = undefined;
		if ((lT || {}).Readable) aT(lT.Readable);
		try {
			er = undefined
		} catch(ah) {}
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
