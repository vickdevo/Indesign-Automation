var doc = app.activeDocument,
    _pages = doc.pages, i, j, k, l,
    _textframes, _tables, _frame, _tables1, _row, _row1, _cell, _cell1, splitstr, newsplit1, rownum, str = " ", str1=" ", content = "", docStory, docTable, myText, myTable, myFrame, _tfheight, currentheight, TabCells, finalheightofcell, newRows, newCells, tempContents="", sttr="", spllitedstr=" ", newsplit=" ";
var mySourceTable = doc.selection[0];  
var temp=0;
for (i = 0; i < _pages.length; i++) {
    content="";
    //alert("Page :"+i);
        if(i==0 || i==1 || i%2!=0||i==(_pages.length-1))
        {
            _tables = doc.stories.item(i).tables;
            for (j = 0; j < _tables.length; j++) {
             //alert ("Page :"+i+", Table :"+j);
            _row = _tables.item(j).rows;                               
            rowlen = _row.length;
                for (k = 4; k < _row.length; k++) {
                 _cell = _row.item(k).cells;
                    for (l = 0; l < _cell.length; l++) {
                        if (k == 4) {
                        str = _cell.item(l).contents;
                        spliitedstr = str.split("M");                  
                        newsplit = str.split(" ");
                        content += newsplit[newsplit.length-1];
                        content += " TemplateOne| |"+newsplit[1]+" "+spliitedstr[0];
                        //alert(content);
                        }
                            else if (k==5) {
                                if(l==0)
                            {
                                str=_cell.item(l).contents;
                                content+=str+"|"+"What does this mean for you?"+"|";
                            }
                            else if((l==1||l==2)){
                                if(l==1)
                            {
                                str=_cell.item(l).contents;
                                content+=str+"|";
                            }
                            else if(l==2)
                            {   
                                str=_cell.item(l).contents;
                                content+=str;
                            }
                                }
                                    }   
                            else {
                                if(l==0){
                                    tempContents=_cell.item(l).contents;
                                    temp=tempContents.indexOf("(", 0);
                                    tempContents=tempContents.replace(")", " ");
                                    if(temp==-1)
                                        {
                                         spllitedstr=tempContents.split("(");
                                         content+=spllitedstr[0]+"|";
                                        }
                             else {
                                spllitedstr=tempContents.split("(");
                                sttr=spllitedstr[1];
                                content+=spllitedstr[0]+"|"+sttr;
                                }
                            }
                            else{
                                    str=_cell.item(l).contents;
                                    content+="|"+str;
                               }
                            }
                        }
                            content+="#";                                        
                    }
                           
                }
            //alert(content);
            _tables.item(0).remove();
             
                myFrame=_pages.item(i).textFrames.add();
                myFrame.geometricBounds=["3p0","3p0","63p0","48p0"];
                myFrame.contents=content;
                myFrame.fit(FitOptions.CONTENT_TO_FRAME);
                myText=myFrame.parentStory.characters.itemByRange(-1, -content.length);
                myTable=myText.convertToTable("|", "#", 4);
                //myTable.appliedTableStyle="Tab1";
                //myTable.cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="ParA";    
                   }
                    else
                    {
                        _tables = doc.stories.item(i).tables;
                        _tables1 = doc.stories.item(i+1).tables;
                        //alert ("Page :"+i+", Table :"+j);
                            _row = _tables.item(0).rows;  
                            _row1 = _tables1.item(0).rows;  
                        // alert("_row length: "+_row.length+"_row1.length: "+_row1.length);
                                for (k = 4; k < _row.length; k++) {
                                    _cell = _row.item(k).cells;
                                    _cell1=_row1.item(k).cells;
                                for (l = 0; l < _cell.length; l++) {
                                    if(_cell.length==2 || _cell.length==1)
                                        {
                                            if (k == 4) {
                                            str = _cell.item(l).contents;
                                            spliitedstr = str.split("M"); 
                                            str1=_cell1.item(l).contents;
                                            //newsplit = str.split(" ");
                                            splitstr=str1.split("M");
                                            //newsplit1=str.split(" ");
                                            //content += newsplit[newsplit.length-1];
                                            content += " TemplateOne#Continued#"+spliitedstr[0]+"|" +"     "+"|"+splitstr[0]+"#";
                                            //alert(content);
                                        }
                                    else if(k==20 || k==26 || k==29){
                                        content+="|    |    |    #";
                                        }
                                    else {
                                        if(l==0){
                                                    ;
                                                }
                                    else {
                                            str=_cell.item(l).contents;
                                            //alert("str:"+str);
                                            str1=_cell.item(l).contents;
                                            //alert("str1:"+str1);
                                            content+=str+"|"+"    "+"|"+str1+"    "+"#";
                                            }
                                        }   
                                    }
                                    else
                                       {
                                            if(l==1)
                                        {
                                            str=_cell.item(l).contents;
                                            str1=_cell.item(l+1).contents;
                                            str2=_cell1.item(l).contents;
                                            str3=_cell1.item(l+1).contents;
                                            content+=str+"|"+str1+"|"+str2+"|"+str3+"#";
                                        }
                                    else
                                        {
                                            ;
                                        }
                                      }
                                        //alert("K:"+k);
                                        //content+="#";                                        
                                    }
                                }
                i=i+1;
                        _tables.item(0).remove();
                _tables1.item(0).remove();
                //alert(content);
                myFrame=_pages.item(i-1).textFrames.add();
                myFrame.geometricBounds=["3p0","54p0","63p0","99p0"];
                myFrame.contents=content;
                myFrame.fit(FitOptions.CONTENT_TO_FRAME);
                myText=myFrame.parentStory.characters.itemByRange(-1, -content.length);
                myTable=myText.convertToTable("|", "#", 4);
                //myTable.appliedTableStyle="Tab1";
                //myTable.cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="ParA";    

    }
                 //alert(content);
                
}


//alert(_pages.length);
    for(i=0;i<_pages.length-1;i++){
        //alert("page number:"+i);
                _frame=_pages.item(i).textFrames;
                //alert("number of frames in that page:"+_frame.length);
          
                //alert("frame number:"+0);
                _tables=_frame.item(0).tables;
                //alert("number of tables in that frame:"+_tables.length);
                _tables.item(0).appliedTableStyle="Plan Design (2 column plans)";
                _rows=_tables.item(0).rows;
                //alert("number of rows in that table:"+_rows.length);
                _tables.item(0).rows.item(0).cells.item(0).merge(_tables.item(0).rows.item(0).cells.item(-1));
                _tables.item(0).rows.item(1).cells.item(0).merge(_tables.item(0).rows.item(1).cells.item(-1));

                for(j=0;j<_rows.length;j++)
                {
                    _cell=_rows.item(j).cells;

                    _rows.item(2).cells.everyItem().appliedCellStyle="Column header name";
                    _rows.item(3).cells.everyItem().appliedCellStyle="Bold Row";
                    
                    for(k=0;k<_cell.length;k++)
                    {
                        
                        if(k==0)
                        {
                            _cell.item(k).paragraphs.everyItem().appliedParagraphStyle="Row header";
 
                         }
                        else if(k==1)
                        {
                            ;
                         }
                     
                        else if(k==2 || k==3)
                        {
                             _cell.item(k).paragraphs.everyItem().appliedParagraphStyle="Table body cell";
                         }
                        else
                        _cell.item(k).appliedCellStyle="BodyCell";
                     }
                }
                _tables.item(0).rows.item(0).cells.item(0).paragraphs.everyItem().appliedParagraphStyle="Plan level headline";
                _tables.item(0).rows.item(1).cells.item(0).paragraphs.everyItem().appliedParagraphStyle="Plan comparison subhead";
     }
