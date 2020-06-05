//DESCRIPTION: Apply languages to frames and styles
//=============================================================
//  Script by Luis Felipe Corullón
//  Contato: lf@corullon.com.br
//  Site: http://lf.corullon.com.br
//=============================================================
/*
MIT License
Copyright (c) 2020 lfcorullon
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
//=============================================================

//IF NO DOCUMENTS ARE OPEN, ALERT AND EXIT SCRIPT EXECUTION
if (app.documents.length == 0 || app.documents[0].visible == false) {
    alert("You must run this script with a document open." , "No document open" , true);
    exit();
    }

//COLECT ALL INDESIGN LANGUAGES
var myLanguages = new Array;
//LOOP THROUGH LANGUAGES
for (var j=app.languagesWithVendors.length-1; j > 0; j--){
    //STORE EACH LANGUAGE NAME IN THE NEW ARRAY
    myLanguages.push(app.languagesWithVendors[j].name);
    }
//SORT THE NEW ARRAY ALPHABETICALLY
myLanguages.sort();

//CREATE A NEW DIALOG WINDOW
var myWindow = new Window ("dialog", "Apply language to styles and texts");
//ADD A STATICTEXT TO THE NEW DIALOG WINDOW
myWindow.add ("statictext", undefined, "Choose the language to apply:");
//ADD A DROPDOWN LIST TO THE NEW DIALOG WINDOW
var myChoice = myWindow.add ("dropdownlist", undefined, myLanguages);
//LOOP THROUGH THE LANGUAGES AND SET THE DEFAULT ONE AS PORTUGUESE
for (var i=0; i<myLanguages.length; i++) {
    if (myLanguages[i] == "Portuguese: Orthographic Agreement") {
        //IF PORTUGUESE: ORTOGRAPHIC AGREEMENT EXISTS TRY TO SET AS SELECTED ITEM
        try {
            myChoice.selection = i;
            }
        //IF PORTUGUESE: ORTOGRAPHIC AGREEMENT DOESN'T EXISTS, SET THE FIRST ONE AS SELECTED ITEM
        catch(e) {
            myChoice.selection = 0;
            }
        }
    }
//ADD NEW GROUP TO THE DIALOG WINDOW
var myButtonGroup = myWindow.add ("group");
//ADD OK BUTTON
var myButtonOK = myButtonGroup.add ("button", undefined, "Confirm", {name:"OK"});
//ADD CANCEL BUTTON
var myButtonCancel = myButtonGroup.add ("button", undefined, "Cancel", {name:"Cancel"});

//SHOW THE DIALOG IN THE USER INTERFACE
var myResult = myWindow.show();

//IF THE USER CLICKS OK
if (myResult != 2) {
    //STORE THE SELECTED LANGUAGE AS THELANGUAGE VARIABLE
    var theLanguage = "$ID/" + myChoice.selection.text;
    
    //STORE PARAGRAPH STYLES, CHARACTER STYLES AND DOCUMENT PAGES
    var paraStyles = app.activeDocument.allParagraphStyles;
    var charStyles = app.activeDocument.allCharacterStyles;
    var pages = app.activeDocument.pages;

//LOOP THROUGH ALL PARAGRAPH STYLES SETTING THE SELECTED LANGUAGE AS THE STYLE LANGUAGE
    for (var j=paraStyles.length-1; j > 0; j--){
        paraStyles[j].appliedLanguage = theLanguage;
        }
//LOOP THROUGH ALL CHARACTER STYLES SETTING THE SELECTED LANGUAGE AS THE STYLE LANGUAGE
    for (var k=charStyles.length-1; k > 0; k--){
        charStyles[k].appliedLanguage = theLanguage;
        }
//LOOP THROUGH ALL PAGES AND ALL TEXT FRAMES SETTING EACH TEXT FRAME CONTENT APPLIED LANGUAGE TO THE SELECTED LANGUAGE
    for (var i = pages.length-1; i >= 0; i--) {
        //LOOP THROUGH TEXT FRAMES
        for (var h = 0; h <= pages[i].textFrames.length-1; h++) {
            app.activeDocument.pages[i].textFrames[h].parentStory.appliedLanguage = theLanguage;
            //IF THE TEXT FRAME CONTAINS TABLES
            if (app.activeDocument.pages[i].textFrames[h].tables.length != 0) {
                //LOOP THROUGH EACH TABLE SETTING THE SELECTED LANGUAGE IN EACH CELL
                for (var q = app.activeDocument.pages[i].textFrames[h].tables.length-1; q >= 0; q--){
                    app.activeDocument.pages[i].textFrames[h].tables[q].cells.everyItem().texts.everyItem().appliedLanguage = theLanguage;
                    }
                }
            }
        }
    //ALERT WHEN EVERYTHING RUNS OK
    alert ("Done!");
    }
