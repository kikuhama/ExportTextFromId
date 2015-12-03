/*
  InDesignデータから，テキスト構造をXMLとして書き出す
  Version 1.0.2

  Changes
  1.0.0: initial version
  1.0.1: ルビをサポート (2015/10/1)
  1.0.2: 複数ファイルをいっぺんに処理したとき，XMLを一つにまとめるよう変更
*/
#target indesign

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var CHARSTYLE_NONE = "[なし]";

function dumpObj(obj) {
    $.writeln("----");
    $.writeln(obj.toString());
    for(var prop in obj) {
	try {
	    $.writeln("name: " + prop + "; value: " + obj[prop]);
	}
	catch(e) {
	    $.writeln("name: " + prop + "; cannot access this property.");
	}
    }
}

function dumpObjToFile(obj, filename) {
    var file = new File(filename);
    if(file && file.open("w")) {
	file.encoding = "utf-8";
	file.writeln(obj.toString());
	for(var prop in obj) {
	    try {
		file.writeln("name: " + prop + "; value: " + obj[prop]);
	    }
	    catch(e) {
		file.writeln("name: " + prop + "; cannot access this property.");
	    }
	}
	file.close();
    }    
}

function selectIdFiles(folder) {
    var idFiles = [];
    var files = folder.getFiles("*");
    for(var i=0; i<files.length; ++i) {
	var f = files[i];
	if(f instanceof Folder) {
	    idFiles = idFiles.concat(selectIdFiles(f));
	}
	else if(f.name.match(/\.indd$/)) {
	    idFiles.push(f);
	}
    }
    return idFiles;
}

function PdfVisibleLayers(pdf) {
    var layers = pdf.graphicLayerOptions.graphicLayers;
    var visibleLayers = "";
    for(var i=0; i<layers.length; ++i) {
	var layer = layers[i];
	if(layer.currentVisibility) {
	    visibleLayers += layer.name;
	}
    }
    return visibleLayers;
}

function TextContents() {
    this.page = 0;
    this.top = 0;
    this.left = 0;
    this.paraStyle = "";
    this.contents = [];
}

TextContents.prototype.setPage = function(page) {
    page_i = parseInt(page);
    if(isNaN(page_i)) {
	this.page = page
    }
    else {
	this.page = page_i;
    }
}

TextContents.prototype.setTop = function(top) {
    top = parseFloat(top);
    if(isNaN(top)) {
	top = 0;
    }
    this.top = top;
}

TextContents.prototype.setLeft = function(left) {
    left = parseFloat(left);
    if(isNaN(left)) {
	left = 0;
    }
    this.left = left;
}

TextContents.prototype.addText = function(t) {
    this.contents.push({
	type: "TEXT",
	text: t
    });
}

TextContents.prototype.addSpecialChar = function(c) {
    this.contents.push({
	type: "SPECIALCHAR",
	text: c
    });
}

TextContents.prototype.addAnchorText = function(at) {
    this.contents.push({
	type: "ANCHORTEXT",
	text: at
    });
}

TextContents.prototype.addRuby = function(text, ruby) {
    this.contents.push({
	type: "RUBY",
	text: text,
	ruby: ruby
    });
}

TextContents.prototype.addGraphic = function(g) {
    var link = g.itemLink;
    this.contents.push({
	type: "GRAPHIC",
	text: link.name,
	layer: PdfVisibleLayers(g)
    });
}

TextContents.prototype.addCharStyle = function(cs) {
    this.contents.push({
	type: "CHARSTYLE",
	text: cs
    });
}

		       
function Analyzer() {
}

Analyzer.prototype.analyzeDoc = function(doc) {
    this.stories = [];
    for(var storyIndex=0; storyIndex<doc.stories.length; ++storyIndex) {
	var story = doc.stories[storyIndex];
	var storyData = this.analyzeStory(story)
	if(storyData && storyData.page != 0) {
	    this.stories.push(storyData);
	}
    }
}

Analyzer.prototype.analyzeStory = function(story) {
    var storyData = {
	paragraphs: [],
	page: 0,
	top: 0
    };
    if(story && story.isValid) {
	for(var paraIndex=0; paraIndex<story.paragraphs.length; ++paraIndex) {
	    var text = this.analyzeParagraph(story.paragraphs[paraIndex]);
	    if(text) {
		storyData.paragraphs.push(text);
	    }
	}
	if(storyData.paragraphs.length > 0) {
	    storyData.page = storyData.paragraphs[0].page;
	    storyData.top = storyData.paragraphs[0].top;
	}
	return storyData;
    }
    return null;
}

Analyzer.prototype.getChild = function(item) {
    if(item.allPageItems.length > 0) {
	return this.getChild(item.allPageItems[0]);
    }
    else {
	return item;
    }
}

Analyzer.prototype.analyzeParagraph = function(paragraph) {
    if(paragraph.isValid) {
	var text = new TextContents();
	var textFrame = paragraph.parentTextFrames[0];
	if(textFrame && textFrame.isValid
	   && textFrame.parentPage && textFrame.parentPage.isValid) {
	    var parent = textFrame.parent;
	    if(!(parent instanceof(Character))) {
		text.setPage(textFrame.parentPage.name);
		text.setTop(textFrame.visibleBounds[0]);
		text.setLeft(textFrame.visibleBounds[1]);
		text.paraStyle = paragraph.appliedParagraphStyle.name;
		var s = "";
		var chStyle = null;
		var ruby = null;
		for(var charIndex=0; charIndex<paragraph.characters.length; charIndex++) {
		    var ch = paragraph.characters[charIndex];
		    if(ch.appliedCharacterStyle.isValid
		       && ch.appliedCharacterStyle.name != chStyle) {
			chStyle = ch.appliedCharacterStyle.name;
			text.addText(s);
			s = "";
			text.addCharStyle(chStyle);
		    }
		    if(ch.allPageItems.length > 0) {
			// anchor object
			text.addText(s);
			s = "";
			var anchorObject = this.getChild(ch);
			if(anchorObject instanceof(TextFrame)) {
			    text.addAnchorText(anchorObject.contents);
			}
			else if(anchorObject instanceof(Graphic)
				|| anchorObject instanceof(PDF)) {
			    text.addGraphic(anchorObject);
			}
		    }
		    else {
			var ch_s = ch.contents.toString();
			// ルビ
			if(!ruby && ch.rubyFlag) {
			    // ruby started
			    text.addText(s);
			    s = "";
			    ruby = ch.rubyString;
			}
			else if(ruby && !ch.rubyFlag) {
			    // ruby finished
			    text.addRuby(s, ruby);
			    s = "";
			    ruby = null;
			}
			else if(ruby && ch.rubyFlag && ruby != ch.rubyString) {
			    // ruby restarted
			    text.addRuby(s, ruby);
			    s = "";
			    ruby = ch.rubyString;
			}
			
			if(ch_s.length > 1) {
			    // special character
			    switch(ch_s) {
			    case "SINGLE_RIGHT_QUOTE":
				s += "’";
				break;
			    case "SINGLE_LEFT_QUOTE":
				s += "‘";
				break;
			    case "DOUBLE_RIGHT_QUOTE":
				s += "”";
				break;
			    case "DOUBLE_LEFT_QUOTE":
				s += "“";
				break;
			    default:
				text.addText(s);
				s = "";
				text.addSpecialChar(ch_s);
			    }
			}
			else {
			    // normal character
			    if(ch_s.match(/\t/)) {
				text.addText(s);
				s = "";
				text.addSpecialChar("TAB");
			    }
			    else {
				s += ch_s;
			    }
			}
		    }
		}
		text.addText(s);
		return text;
	    }
	}
    }
}

Analyzer.prototype.xmlText = function(s) {
    try {
	var s2 = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
	s2 = s2.replace(/[\u0000-\u001f]/, "");
	return s2;
    }
    catch(e) {
	return s;
    }
}

Analyzer.prototype.makeParagraphEle = function(para) {
    var charstyleStarted = false;
    var paraEle = new XML("<p/>");
    var contentEle = paraEle;
    paraEle.@style = para.paraStyle;
    for(var contentIndex=0; contentIndex<para.contents.length; ++contentIndex) {
	var content = para.contents[contentIndex];
	var childEle;
	switch(content.type) {
	case "TEXT":
	    contentEle.appendChild(content.text);
	    break;
	case "SPECIALCHAR":
	    childEle = new XML("<specialChar/>");
	    childEle.@name = content.text;
	    contentEle.appendChild(childEle);
	    break;
	case "ANCHORTEXT":
	    childEle = new XML("<anchorText/>");
	    childEle.appendChild(content.text);
	    contentEle.appendChild(childEle);
	    break;
	case "GRAPHIC":
	    childEle = new XML("<anchorGraphic/>");
	    childEle.@name = content.text;
	    childEle.@layer = content.layer;
	    contentEle.appendChild(childEle);
	    break;
	case "CHARSTYLE":
	    if(content.text != CHARSTYLE_NONE) {
		contentEle = new XML("<c/>");
		charstyleStarted = true;
		contentEle.@style = content.text;
		paraEle.appendChild(contentEle);
	    }
	    else {
		// 文字スタイル解除
		contentEle = paraEle;
	    }
	    break;
	case "RUBY":
	    var rubyEle = new XML("<ruby/>");
	    rubyEle.appendChild(content.text);
	    var rtEle = new XML("<rt/>");
	    rtEle.appendChild(content.ruby);
	    rubyEle.appendChild(rtEle);
	    contentEle.appendChild(rubyEle);
	    break;
	}
    }

    return paraEle;
}

Analyzer.prototype.xml = function() {
    this.stories.sort(function(a, b) {
	if(a.page != b.page) {
	    var pa = Number(a.page);
	    var pb = Number(b.page);
	    if(isNaN(pa) && isNaN(pb)) {
		// a.page, b.pageともに文字列
		if(String(a.page) > String(b.page)) {
		    return 1;
		}
		else {
		    return -1;
		}
	    }
	    else if(isNaN(pa)) {
		return 1;
	    }
	    else if(isNaN(pb)) {
		return -1;
	    }
	    else {
		// a.page, b.pageともに数値
		return pa - pb;
	    }
	}
	else{
	    return a.top - b.top;
	}
    });
    var xmldoc = new XML("<pages/>");
    var currentPage = null;
    var pageEle;
    var storyEle;
    for(var storyIndex=0; storyIndex<this.stories.length; ++storyIndex) {
	var story = this.stories[storyIndex];
	if(story.page != currentPage) {
	    pageEle = new XML("<page/>");
	    pageEle.@number = story.page;
	    xmldoc.appendChild(pageEle);
	    currentPage = story.page;
	}
	storyEle = new XML("<story/>");
	for(var paraIndex=0; paraIndex<story.paragraphs.length; ++paraIndex) {
	    var para = story.paragraphs[paraIndex];
	    storyEle.appendChild(this.makeParagraphEle(para));
	}
	pageEle.appendChild(storyEle);
    }

    return xmldoc;
}

var getSaveFile = function() {
    var file = new File("export_text.xml");
    file = file.saveDlg("エクスポートファイル名を指定してください。",
			"*.xml",
			false);
    return file;
}



var analyzer = new Analyzer();
XML.prettyPrinting = true;
var xmldoc = new XML("<files/>");
var canceled = false;

if(app.documents.length > 0) {
    analyzer.analyzeDoc(app.activeDocument);
    var fileEle = new XML("<file/>");
    fileEle.@name = app.activeDocument.name;
    fileEle.appendChild(analyzer.xml());
    xmldoc.appendChild(fileEle);
}
else {
    var folderName = Folder.selectDialog("フォルダを選択してください。");
    if(folderName) {
	var folder = new Folder(folderName);
	var idFiles = selectIdFiles(folder);
	for(var i=0; i<idFiles.length; ++i) {
	    var doc = app.open(idFiles[i]);
	    analyzer.analyzeDoc(doc);
	    doc.close(SaveOptions.NO);
	    var fileEle = new XML("<file/>");
	    fileEle.@name = idFiles[i];
	    fileEle.appendChild(analyzer.xml());
	    xmldoc.appendChild(fileEle);
	}
    }
    else {
	canceled = true;
    }
}

if(!canceled) {
    var file = getSaveFile();
    if(file && file.open("w")) {
	file.encoding = "utf-8";
	file.write("<?xml version='1.0' encoding='utf-8'?>");
	file.write(xmldoc.toXMLString());
	file.close();
    }
}

alert("finished");
