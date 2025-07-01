export class Theme {
	/** @type {string} */ name;
	/**
	 * @param {Element} theme 
	 */
	constructor(theme) {
		const name = theme.getAttribute('name');
		if (name) this.name = name;
		const custClrLst = theme.querySelector('custClrList');
		if (custClrLst) this.customColorList = custClrLst;
		const extraClrSchemeLst = theme.querySelector('extraClrSchemeLst');
		if (extraClrSchemeLst) this.extraColorSchemeList = extraClrSchemeLst;
		const themeElements = theme.querySelector('themeElements');
		if (themeElements) this.themeElements = ThemeElements.from(themeElements);
	}
}

export class ThemeElements {
	/** @type {ColorScheme?} */ colorScheme = null;
	/** @type {FormatScheme?} */ formatScheme = null;
	/** @type {FontScheme?} */ fontScheme = null;
	/**
	 * @param {Element} themeElements 
	 */
	static from(themeElements) {
		const self = new ThemeElements();
		const clrScheme = themeElements.querySelector('clrScheme');
		if (clrScheme) self.colorScheme = ColorScheme.from(clrScheme);
		const fmtScheme = themeElements.querySelector('fmtScheme');
		if (fmtScheme) self.formatScheme = FormatScheme.from(fmtScheme);
		const fontScheme = themeElements.querySelector('fontScheme');
		if (fontScheme) self.fontScheme = FontScheme.from(fontScheme);
		return self;
	}
}

export class ColorScheme {
	/** @type {string?} */ name = null;
	/** @type {ColorSchemeChild?} */ accent1 = null;
	/** @type {ColorSchemeChild?} */ accent2 = null;
	/** @type {ColorSchemeChild?} */ accent3 = null;
	/** @type {ColorSchemeChild?} */ accent4 = null;
	/** @type {ColorSchemeChild?} */ accent5 = null;
	/** @type {ColorSchemeChild?} */ accent6 = null;
	/** @type {ColorSchemeChild?} */ bk1 = null;
	/** @type {ColorSchemeChild?} */ bk2 = null;
	/** @type {ColorSchemeChild?} */ dk1 = null;
	/** @type {ColorSchemeChild?} */ dk2 = null;
	/** @type {ColorSchemeChild?} */ folHlink = null;
	/** @type {ColorSchemeChild?} */ hlink = null;
	/** @type {ColorSchemeChild?} */ lt1 = null;
	/** @type {ColorSchemeChild?} */ lt2 = null;
	/** @type {ColorSchemeChild?} */ phClr = null;
	/** @type {ColorSchemeChild?} */ tx1 = null;
	/** @type {ColorSchemeChild?} */ tx2 = null;
	/**
	 * @param {Element} clrScheme 
	 */
	static from(clrScheme) {
		const self = new ColorScheme();
		self.name = clrScheme.getAttribute('name') || '';
		const accent1 = clrScheme.querySelector('accent1 > *');
		if (accent1) self.accent1 = self.getColor(accent1);
		const accent2 = clrScheme.querySelector('accent2 > *');
		if (accent2) self.accent2 = self.getColor(accent2);
		const accent3 = clrScheme.querySelector('accent3 > *');
		if (accent3) self.accent3 = self.getColor(accent3);
		const accent4 = clrScheme.querySelector('accent4 > *');
		if (accent4) self.accent4 = self.getColor(accent4);
		const accent5 = clrScheme.querySelector('accent5 > *');
		if (accent5) self.accent5 = self.getColor(accent5);
		const accent6 = clrScheme.querySelector('accent6 > *');
		if (accent6) self.accent6 = self.getColor(accent6);
		const bk1 = clrScheme.querySelector('bk1 > *');
		if (bk1) self.bk1 = self.getColor(bk1);
		const bk2 = clrScheme.querySelector('bk2 > *');
		if (bk2) self.bk2 = self.getColor(bk2);
		const dk1 = clrScheme.querySelector('dk1 > *');
		if (dk1) self.dk1 = self.getColor(dk1);
		const dk2 = clrScheme.querySelector('dk2 > *');
		if (dk2) self.dk2 = self.getColor(dk2);
		const folHlink = clrScheme.querySelector('folHlink > *');
		if (folHlink) self.folHlink = self.getColor(folHlink);
		const hlink = clrScheme.querySelector('hlink > *');
		if (hlink) self.hlink = self.getColor(hlink);
		const lt1 = clrScheme.querySelector('lt1 > *');
		if (lt1) self.lt1 = self.getColor(lt1);
		const lt2 = clrScheme.querySelector('lt2 > *');
		if (lt2) self.lt2 = self.getColor(lt2);
		const phClr = clrScheme.querySelector('phClr > *');
		if (phClr) self.phClr = self.getColor(phClr);
		const tx1 = clrScheme.querySelector('tx1 > *');
		if (tx1) self.tx1 = self.getColor(tx1);
		const tx2 = clrScheme.querySelector('tx2 > *');
		if (tx2) self.tx2 = self.getColor(tx2);
		return self;
	}

	/**
	 * @param {Element} child 
	 * @returns {ColorSchemeChild?}
	 */
	getColor(child) {
		switch (child.localName) {
			case 'hslClr':
				return HslColor.from(child);
			case 'prstClr':
				return PresetColor.from(child);
			// case 'schemeClr':
			case 'scrgbClr':
				return RgbColorModelPercentage.from(child);
			case 'srgbClr':
				return RgbColorModelHex.from(child);
			case 'sysClr':
				return SystemColor.from(child);
			default:
				return null;
		}
	}
}

/**
 * @typedef {object} ColorSchemeChild
 * @prop {() => string} toString
 */

/**
 * @implements {ColorSchemeChild}
 */
export class HslColor {
	/** @type {string?} */ hue = null;
	/** @type {string?} */ sat = null;
	/** @type {string?} */ lum = null;

	/**
	 * @param {Element} hslClr 
	 */
	static from(hslClr) {
		const self = new HslColor();
		const hue = hslClr.getAttribute('hue');
		if (hue) self.hue = hue;
		const sat = hslClr.getAttribute('sat');
		if (sat) self.sat = sat;
		const lum = hslClr.getAttribute('lum');
		if (lum) self.lum = lum;
		return self;
	}

	toString() {
		return this.hue ? `hsl(${Number.parseInt(this.hue) / 6000}deg ${this.sat} ${this.lum})` : '';
	}
}

/**
 * @implements {ColorSchemeChild}
 */
export class PresetColor {
	static Values = {
		aliceBlue: [240, 248, 255],
		antiqueWhite: [250, 235, 215],
		aqua: [0, 255, 255],
		aquamarine: [127, 255, 212],
		azure: [240, 255, 255],
		beige: [245, 245, 220],
		bisque: [255, 228, 196],
		black: [0, 0, 0],
		blanchedAlmond: [255, 235, 205],
		blue: [0, 0, 255],
		blueViolet: [138, 43, 226],
		brown: [165, 42, 42],
		burlyWood: [222, 184, 135],
		cadetBlue: [95, 158, 160],
		chartreuse: [127, 255, 0],
		chocolate: [210, 105, 30],
		coral: [255, 127, 80],
		cornflowerBlue: [100, 149, 237],
		cornsilk: [255, 248, 220],
		crimson: [220, 20, 60],
		cyan: [0, 255, 255],
		darkBlue: [0, 0, 139],
		darkCyan: [0, 139, 139],
		darkGoldenrod: [184, 134, 11],
		darkGray: [169, 169, 169],
		darkGreen: [0, 100, 0],
		darkGrey: [169, 169, 169],
		darkKhaki: [189, 183, 107],
		darkMagenta: [139, 0, 139],
		darkOliveGreen: [85, 107, 47],
		darkOrange: [255, 140, 0],
		darkOrchid: [153, 50, 204],
		darkRed: [139, 0, 0],
		darkSalmon: [233, 150, 122],
		darkSeaGreen: [143, 188, 143],
		darkSlateBlue: [72, 61, 139],
		darkSlateGray: [47, 79, 79],
		darkSlateGrey: [47, 79, 79],
		darkTurquoise: [0, 206, 209],
		darkViolet: [148, 0, 211],
		deepPink: [255, 20, 147],
		deepSkyBlue: [0, 191, 255],
		dimGray: [105, 105, 105],
		dimGrey: [105, 105, 105],
		dkBlue: [0, 0, 139],
		dkCyan: [0, 139, 139],
		dkGoldenrod: [184, 134, 11],
		dkGray: [169, 169, 169],
		dkGreen: [0, 100, 0],
		dkGrey: [169, 169, 169],
		dkKhaki: [189, 183, 107],
		dkMagenta: [139, 0, 139],
		dkOliveGreen: [85, 107, 47],
		dkOrange: [255, 140, 0],
		dkOrchid: [153, 50, 204],
		dkRed: [139, 0, 0],
		dkSalmon: [233, 150, 122],
		dkSeaGreen: [143, 188, 143],
		dkSlateBlue: [72, 61, 139],
		dkSlateGray: [47, 79, 79],
		dkSlateGrey: [47, 79, 79],
		dkTurquoise: [0, 206, 209],
		dkViolet: [148, 0, 211],
		dodgerBlue: [30, 144, 255],
		firebrick: [178, 34, 34],
		floralWhite: [255, 250, 240],
		forestGreen: [34, 139, 34],
		fuchsia: [255, 0, 255],
		gainsboro: [220, 220, 220],
		ghostWhite: [248, 248, 255],
		gold: [255, 215, 0],
		goldenrod: [218, 165, 32],
		gray: [128, 128, 128],
		green: [0, 128, 0],
		greenYellow: [173, 255, 47],
		grey: [128, 128, 128],
		honeydew: [240, 255, 240],
		hotPink: [255, 105, 180],
		indianRed: [205, 92, 92],
		indigo: [75, 0, 130],
		ivory: [255, 255, 240],
		khaki: [240, 230, 140],
		lavender: [230, 230, 250],
		lavenderBlush: [255, 240, 245],
		lawnGreen: [124, 252, 0],
		lemonChiffon: [255, 250, 205],
		lightBlue: [173, 216, 230],
		lightCoral: [240, 128, 128],
		lightCyan: [224, 255, 255],
		lightGoldenrodYellow: [250, 250, 210],
		lightGray: [211, 211, 211],
		lightGreen: [144, 238, 144],
		lightGrey: [211, 211, 211],
		lightPink: [255, 182, 193],
		lightSalmon: [255, 160, 122],
		lightSeaGreen: [32, 178, 170],
		lightSkyBlue: [135, 206, 250],
		lightSlateGray: [119, 136, 153],
		lightSlateGrey: [119, 136, 153],
		lightSteelBlue: [176, 196, 222],
		lightYellow: [255, 255, 224],
		lime: [0, 255, 0],
		limeGreen: [50, 255, 50],
		linen: [250, 240, 230],
		ltBlue: [173, 216, 230],
		ltCoral: [240, 128, 128],
		ltCyan: [224, 255, 255],
		ltGoldenrodYellow: [250, 250, 210],
		ltGray: [211, 211, 211],
		ltGreen: [144, 238, 144],
		ltGrey: [211, 211, 211],
		ltPink: [255, 182, 193],
		ltSalmon: [255, 160, 122],
		ltSeaGreen: [32, 178, 170],
		ltSkyBlue: [135, 206, 250],
		ltSlateGray: [119, 136, 153],
		ltSlateGrey: [119, 136, 153],
		ltSteelBlue: [176, 196, 222],
		ltYellow: [255, 255, 224],
		magenta: [255, 0, 255],
		maroon: [128, 0, 0],
		medAquamarine: [102, 205, 170],
		medBlue: [0, 0, 205],
		mediumAquamarine: [102, 205, 170],
		mediumBlue: [0, 0, 205],
		mediumOrchid: [186, 85, 211],
		mediumPurple: [147, 112, 219],
		mediumSeaGreen: [60, 179, 113],
		mediumSlateBlue: [123, 104, 238],
		mediumSpringGreen: [0, 250, 154],
		mediumTurquoise: [72, 209, 204],
		mediumVioletRed: [199, 21, 133],
		medOrchid: [186, 85, 211],
		medPurple: [147, 112, 219],
		medSeaGreen: [60, 179, 113],
		medSlateBlue: [123, 104, 238],
		medSpringGreen: [0, 250, 154],
		medTurquoise: [72, 209, 204],
		medVioletRed: [199, 21, 133],
		midnightBlue: [25, 25, 112],
		mintCream: [245, 255, 250],
		mistyRose: [255, 228, 225],
		moccasin: [255, 228, 181],
		navajoWhite: [255, 222, 173],
		navy: [0, 0, 128],
		oldLace: [253, 245, 230],
		olive: [128, 128, 0],
		oliveDrab: [107, 142, 35],
		orange: [255, 165, 0],
		orangeRed: [255, 69, 0],
		orchid: [218, 112, 214],
		paleGoldenrod: [238, 232, 170],
		paleGreen: [152, 251, 152],
		paleTurquoise: [175, 238, 238],
		paleVioletRed: [219, 112, 147],
		papayaWhip: [255, 239, 213],
		peachPuff: [255, 218, 185],
		peru: [205, 133, 63],
		pink: [255, 192, 203],
		plum: [221, 160, 221],
		powderBlue: [176, 224, 230],
		purple: [128, 0, 128],
		red: [255, 0, 0],
		rosyBrown: [188, 143, 143],
		royalBlue: [65, 105, 225],
		saddleBrown: [139, 69, 19],
		salmon: [250, 128, 114],
		sandyBrown: [244, 164, 96],
		seaGreen: [46, 139, 87],
		seaShell: [255, 245, 238],
		sienna: [160, 82, 45],
		silver: [192, 192, 192],
		skyBlue: [135, 206, 235],
		slateBlue: [106, 90, 205],
		slateGray: [112, 128, 144],
		slateGrey: [112, 128, 144],
		snow: [255, 250, 250],
		springGreen: [0, 255, 127],
		steelBlue: [70, 130, 180],
		tan: [210, 180, 140],
		teal: [0, 128, 128],
		thistle: [216, 191, 216],
		tomato: [255, 99, 71],
		turquoise: [64, 224, 208],
		violet: [238, 130, 238],
		wheat: [245, 222, 179],
		white: [255, 255, 255],
		whiteSmoke: [245, 245, 245],
		yellow: [255, 255, 0],
		yellowGreen: [154, 205, 50],
	};
	/** @type {(keyof PresetColor.Values)?} */ val = null;
	/**
	 * @param {Element} prstClr 
	 */
	static from(prstClr) {
		const self = new PresetColor();
		const val = prstClr.getAttribute('val');
		// @ts-ignore
		if (val) self.val = val;
		return self;
	}

	toString() {
		const rgb = PresetColor.Values[this.val];
		return `rgb(${rgb.join(' ')})`;
	}
}

/**
 * @implements {ColorSchemeChild}
 */
export class RgbColorModelHex {
	/** @type {string?} */ val = null;
	/** @type {number} */ r = NaN;
	/** @type {number} */ g = NaN;
	/** @type {number} */ b = NaN;
	/**
	 * 
	 * @param {Element} srgbClr 
	 */
	static from(srgbClr) {
		const self = new RgbColorModelHex();
		const val = srgbClr.getAttribute('val');
		if (val) {
			self.val = val;
			self.r = Number.parseInt(val.substring(0, 2), 16);
			self.g = Number.parseInt(val.substring(2, 4), 16);
			self.b = Number.parseInt(val.substring(4, 6), 16);
		}
		return self;
	}

	toString() {
		return '#' + this.val;
	}
}

/**
 * @implements {ColorSchemeChild}
 */
export class RgbColorModelPercentage {
	/** @type {number} */ r = NaN;
	/** @type {number} */ g = NaN;
	/** @type {number} */ b = NaN;
	/**
	 * @param {Element} scrgbClr 
	 */
	static from(scrgbClr) {
		const self = new RgbColorModelPercentage();
		const r = scrgbClr.getAttribute('r');
		if (r) self.r = Number.parseInt(r);
		const g = scrgbClr.getAttribute('g');
		if (g) self.g = Number.parseInt(g);
		const b = scrgbClr.getAttribute('b');
		if (b) self.g = Number.parseInt(b);
		return self;
	}
}

/**
 * @implements {ColorSchemeChild}
 */
export class SystemColor {
	static Values = {
		/** @deprecated */ '3dDkShadow': 'ThreeDDarkShadow',
		/** @deprecated */ '3dLight': 'ThreeDFace',
		/** @deprecated */ activeBorder: 'ActiveBorder',
		/** @deprecated */ activeCaption : 'ActiveCaption',
		/** @deprecated */ appWorkspace: 'AppWorkspace',
		/** @deprecated */ background: 'Background',
		btnFace: 'ButtonFace',
		/** @deprecated */ btnHighlight: 'ButtonHighlight',
		/** @deprecated */ btnShadow: 'ButtonShadow',
		btnText: 'ButtonText',
		/** @deprecated */ captionText: 'CaptionText',
		grayText: 'GrayText',
		highlight: 'Highlight',
		highlightText: 'HighlightText',
		hotLight: 'LinkText',
		/** @deprecated */ inactiveBorder: 'InactiveBorder',
		/** @deprecated */ inactiveCaption: 'InactiveCaption',
		/** @deprecated */ inactiveCaptionText: 'InactiveCaptionText',
		/** @deprecated */ infoBk: 'InfoBackground',
		/** @deprecated */ infoText: 'InfoText',
		/** @deprecated */ menu: 'Menu',
		/** @deprecated */ menuText: 'MenuText',
		/** @deprecated */ scrollBar: 'ScrollBar',
		/** @deprecated */ window: 'Window',
		/** @deprecated */ windowFrame: 'WindowFrame',
		/** @deprecated */ windowText: 'WindowText',
	};
	/** @type {(keyof SystemColor.Values)?} */ val = null;
	/** @type {string?} */ lastColor = null;

	/**
	 * @param {Element} sysClr 
	 */
	static from(sysClr) {
		const self = new SystemColor();
		/** @type {(keyof SystemColor.Values)?} */ //@ts-ignore
		const val = sysClr.getAttribute('val');
		if (val) self.val = val;
		const lastColor = sysClr.getAttribute('lastClr');
		if (lastColor) self.lastColor = lastColor;
		return self;
	}

	toString() {
		if (this.lastColor) return '#' + this.lastColor;
		else return SystemColor.Values[this.val];
	}
}

export class FormatScheme {
	/**
	 * @param {Element} fmtScheme 
	 */
	static from(fmtScheme) {
		const self = new FormatScheme();
		return self;
	}
}

export class FontScheme {
	/** @type {string?} */ name = null;
	/** @type {FontSchemeChild?} */ major = null;
	/** @type {FontSchemeChild?} */ minor = null;
	/**
	 * @param {Element} fontScheme 
	 */
	static from(fontScheme) {
		const self = new FontScheme();
		self.name = fontScheme.getAttribute('name');
		const major = fontScheme.querySelector('majorFont');
		if (major) self.major = FontSchemeChild.from(major);
		const minor = fontScheme.querySelector('minorFont');
		if (minor) self.minor = FontSchemeChild.from(minor);
		return self;
	}
}

export class FontSchemeChild {
	/** @type {string?} */ latin = null;
	/** @type {string?} */ ea = null;
	/** @type {string?} */ cs = null;
	/** @type {Element[]} */ local = [];

	/**
	 * @param {Element} majorOrMinorFont 
	 */
	static from(majorOrMinorFont) {
		const self = new FontSchemeChild();
		const latin = majorOrMinorFont.querySelector('latin');
		if (latin) self.latin = latin.getAttribute('typeface');
		const ea = majorOrMinorFont.querySelector('ea');
		if (ea) self.ea = ea.getAttribute('typeface');
		const cs = majorOrMinorFont.querySelector('cs');
		if (cs) self.cs = cs.getAttribute('typeface');
		this.local = Array.from(majorOrMinorFont.querySelectorAll('font'));
		return self;
	}
	/**
	 * Gets a typeface of the script.
	 * @param {string} script 
	 */
	getFace(script) {
		switch (script) {
			case 'latin':
				return this.latin;
			case 'ea':
				return this.ea;
			case 'cs':
				return this.cs;
			default:
				const font = this.local.find(f => f.getAttribute('script') === script);
				return font ? font.getAttribute('typeface') : null;
		}
	}

	/**
	 * Sets a typeface of the script.
	 * @param {string} script 
	 * @param {string} typeface 
	 */
	setFace(script, typeface) {
		switch (script) {
			case 'latin':
				this.latin = typeface;
				return;
			case 'ea':
				this.ea = typeface;
				return;
			case 'cs':
				this.cs = typeface;
				return;
			default:
				const font = this.local.find(f => f.getAttribute('script') === script);
				if (font) {
					font.setAttribute('typeface', typeface);
				} else {
					const newFont = /** @type {Element} */ (this.local[0].cloneNode());
					newFont.setAttribute('script', script);
					newFont.setAttribute('typeface', typeface);
					this.local.push(newFont);
				}
		}
	}
}

export class TextAlign {
	static Types = {
		ctr: 'center',
		dist: 'justify',
		just: 'justify',
		justLow: 'justify',
		l: 'left',
		r: 'right',
		thaiDist: 'justify',
	};
}

export class TextAnchoring {
	static Types = {
		b: 'bottom',
		ctr: 'middle',
		// dist: null,
		// just: null,
		t: 'top',
	};
}