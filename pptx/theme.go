/*
MIT License

Copyright (c) 2026 Misael Monterroca

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

package pptx

// defaultTheme is ppt/theme/theme1.xml: a complete, valid Office theme
// (color scheme, font scheme, format scheme). Ported verbatim from
// docxgo's word/theme/theme1.xml — theme1.xml is pure DrawingML (the "a:"
// namespace only) with no WordprocessingML content whatsoever, so nothing
// about it needed to change beyond the part path it is written to.
//
// It is a hand-authored literal rather than a struct tree — same treatment
// as docxgo gave it, and for the same reason: dynamic, brand-driven themes
// are out of scope for the walking skeleton, and modeling the fill/line/
// effect style lists as Go structs buys nothing until something needs to
// vary them.
const defaultTheme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
	<a:themeElements>
		<a:clrScheme name="Office">
			<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
			<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
			<a:dk2><a:srgbClr val="44546A"/></a:dk2>
			<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
			<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
			<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
			<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
			<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
			<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
			<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
			<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
			<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
		</a:clrScheme>
		<a:fontScheme name="Office">
			<a:majorFont>
				<a:latin typeface="Calibri Light"/>
				<a:ea typeface=""/>
				<a:cs typeface=""/>
			</a:majorFont>
			<a:minorFont>
				<a:latin typeface="Calibri"/>
				<a:ea typeface=""/>
				<a:cs typeface=""/>
			</a:minorFont>
		</a:fontScheme>
		<a:fmtScheme name="Office">
			<a:fillStyleLst>
				<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
				<a:gradFill rotWithShape="1">
					<a:gsLst>
						<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
						<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
						<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
					</a:gsLst>
					<a:lin ang="16200000" scaled="1"/>
				</a:gradFill>
				<a:gradFill rotWithShape="1">
					<a:gsLst>
						<a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
						<a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
						<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
					</a:gsLst>
					<a:lin ang="16200000" scaled="1"/>
				</a:gradFill>
			</a:fillStyleLst>
			<a:lnStyleLst>
				<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
				<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
				<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
			</a:lnStyleLst>
			<a:effectStyleLst>
				<a:effectStyle><a:effectLst/></a:effectStyle>
				<a:effectStyle><a:effectLst/></a:effectStyle>
				<a:effectStyle>
					<a:effectLst>
						<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
							<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr>
						</a:outerShdw>
					</a:effectLst>
				</a:effectStyle>
			</a:effectStyleLst>
			<a:bgFillStyleLst>
				<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
				<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>
				<a:gradFill rotWithShape="1">
					<a:gsLst>
						<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>
						<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>
						<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>
					</a:gsLst>
					<a:lin ang="16200000" scaled="1"/>
				</a:gradFill>
			</a:bgFillStyleLst>
		</a:fmtScheme>
	</a:themeElements>
	<a:objectDefaults/>
	<a:extraClrSchemeLst/>
</a:theme>`
