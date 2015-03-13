var doc = app.activeDocument,
    _pages = doc.pages, i, j, k, l,
    _textframes, _tables, _frame, _tables1, _tables2, _rows,  _row, _row1, _row2, _cell, _cell1,_cell2, splitstr, newsplit1, rownum, str = " ", str1=" ", content = "", docStory, docTable, myText, myTable, myFrame, _tfheight, currentheight, TabCells, finalheightofcell, newRows, newCells, tempContents="", sttr="", spllitedstr=" ", newsplit=" ";
var mySourceTable = doc.selection[0];
var temp=0;
for(i=1;i<_pages.length;i++)
{
    _frame=_pages.item(i).textFrames;
    _tables=_frame.item(0).tables;
    _rows=_tables.item(0).rows; 
    app.findTextPreferences=NothingEnum.nothing;
    app.changeTextPreferences=NothingEnum.nothing;
    app.findChangeTextOptions.caseSensitive=true;
    app.findChangeTextOptions.includeMasterPages=true;
    app.findChangeTextOptions.wholeWord=true;
    app.findTextPreferences.findWhat="Not Covered";
    app.changeTextPreferences.changeTo="Not covered";
    app.changeText();
    app.findTextPreferences=NothingEnum.nothing;
    app.changeTextPreferences=NothingEnum.nothing; 
    
    app.findTextPreferences=NothingEnum.nothing;
    app.changeTextPreferences=NothingEnum.nothing;
    app.findChangeTextOptions.caseSensitive=true;
    app.findChangeTextOptions.includeMasterPages=true;
    app.findChangeTextOptions.wholeWord=true;
    app.findTextPreferences.findWhat="Medical";
    app.changeTextPreferences.changeTo="medical";
    app.changeText();
    app.findTextPreferences=NothingEnum.nothing;
    app.changeTextPreferences=NothingEnum.nothing;   
    _tables.item(0).appliedTableStyle="Plan Design (2 column plans)";
    _rows.item(0).cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="Plan level headline";
    _rows.item(0).cells.everyItem().appliedCellStyle="Column Gap";
    _rows.item(1).cells.everyItem().paragraphs.everyItem().leftIndent="0in";
    _rows.item(1).cells.everyItem().leftInset="0in";
    _rows.item(1).cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="Plan comparison subhead";
    _rows.item(2).cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="Plan name";
    _rows.item(2).cells.everyItem().appliedCellStyle="Column header name";
    _rows.item(1).cells.everyItem().height="2p09";
    _rows.item(3).cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="Row header";
    _rows.item(3).cells.everyItem().appliedCellStyle="Bold Row";
    
    if(_tables.item(0).width==45)
    {
      _tables.item(0).rows.item(0).cells.item(-5).merge(_tables.item(0).rows.item(0).cells.item(-1));
      _tables.item(0).rows.item(1).cells.item(-5).merge(_tables.item(0).rows.item(1).cells.item(-1)); 
      _tables.item(0).rows.item(2).cells.item(-2).merge(_tables.item(0).rows.item(2).cells.item(-1));
      _tables.item(0).columns.item(2).width="0p3";
      _tables.item(0).columns.item(0).width="4p105";
      _tables.item(0).columns.item(3).width="9p30";
      _tables.item(0).columns.item(4).width="9p32";
      for(j=4;j<_rows.length;j++)
     {
           _tables.item(0).rows.item(j).cells.item(0).merge(_tables.item(0).rows.item(j).cells.item(1));
           _cell=_rows.item(j).cells;
           for(k=0;k<_cell.length;k++)
           {
                        if(k==0)
                        {
                            app.findGrepPreferences=NothingEnum.nothing;
                            app.changeGrepPreferences=NothingEnum.nothing;
                            app.findGrepPreferences.findWhat="  +";
                            app.changeGrepPreferences.changeTo=" ";
                            app.changeGrep();                          
                            app.findGrepPreferences=NothingEnum.nothing;
                            app.changeGrepPreferences=NothingEnum.nothing;  
                            if(_cell.item(k).contents=="Outpatient surgery \n(Ambulatory Surgical Center/ Hospital)")
                            {
                                _cell.item(k).contents="Outpatient surgery (Ambulatory Surgical Center/Hospital)";
                            }
                            else if(_cell.item(k).contents=="Pediatric eye exam \n(1 visit per year)")
                            {
                                _cell.item(k).contents="Pediatric eye exam (1 visit per year)";   
                            }                            
                            else if(_cell.item(k).contents=="Dental check-up/preventive dental care  (2 visits per year)(2)")
                            {
                                _cell.item(k).contents="Dental check-up/preventive dental care (2 visits per year)(2)";   
                            }                            
                                                                              
                             _cell.item(k).paragraphs.everyItem().appliedParagraphStyle="Row header";
                         }
                        else if(k==1)
                        {
                            ;   
                        }
                        else
                        {   
                            app.findTextPreferences=NothingEnum.nothing;
                            app.changeTextPreferences=NothingEnum.nothing;
                            app.findChangeTextOptions.caseSensitive=true;
                            app.findChangeTextOptions.includeMasterPages=true;
                            app.findChangeTextOptions.wholeWord=true;
                            app.findTextPreferences.findWhat="deductible";
                            app.changeTextPreferences.changeTo="ded";
                            app.changeText();
                            app.findTextPreferences=NothingEnum.nothing;
                            app.changeTextPreferences=NothingEnum.nothing;   
                            _cell.item(k).paragraphs.everyItem().appliedParagraphStyle="Table body cell";
                         }
            }
      }

             _tables.item(0).rows.item(17).cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="Row header";
             _tables.item(0).rows.item(17).cells.everyItem().appliedCellStyle="Bold Row";
             _tables.item(0).rows.item(19).cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="Row header";
             _tables.item(0).rows.item(19).cells.everyItem().appliedCellStyle="Bold Row";  
             _tables.item(0).rows.item(22).cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="Row header";
             _tables.item(0).rows.item(22).cells.everyItem().appliedCellStyle="Bold Row";
             _tables.item(0).columns.item(2).cells.everyItem().appliedCellStyle="Column Gap";

     }
    else if (_tables.item(0).width==96)
    {
    _tables.item(0).rows.item(0).cells.item(-7).merge(_tables.item(0).rows.item(0).cells.item(-11));
    _tables.item(0).rows.item(0).cells.item(-5).merge(_tables.item(0).rows.item(0).cells.item(-1));
    _tables.item(0).rows.item(1).cells.item(-7).merge(_tables.item(0).rows.item(1).cells.item(-11));
    _tables.item(0).rows.item(1).cells.item(-5).merge(_tables.item(0).rows.item(1).cells.item(-1));  
    _tables.item(0).rows.item(2).cells.item(3).merge(_tables.item(0).rows.item(2).cells.item(4)); 
    _tables.item(0).rows.item(2).cells.item(6).merge(_tables.item(0).rows.item(2).cells.item(5)); 
    _tables.item(0).rows.item(2).cells.item(8).merge(_tables.item(0).rows.item(2).cells.item(7));
    _tables.item(0).columns.item(0).width="1.8625in";
    _tables.item(0).columns.item(1).width="1.8625in";
    _tables.item(0).columns.item(2).width="0.05in";
    _tables.item(0).columns.item(5).width="1in";
    _tables.item(0).columns.item(8).width="0.05in";
    _tables.item(0).columns.item(3).width="1.8625in";
    _tables.item(0).columns.item(4).width="1.8625in";
    _tables.item(0).columns.item(6).width="1.8625in";    
    _tables.item(0).columns.item(7).width="1.8625in"; 
    _tables.item(0).columns.item(9).width="1.8625in";
    _tables.item(0).columns.item(10).width="1.8625in";
    for(j=4;j<_rows.length;j++)
     {
           _tables.item(0).rows.item(j).cells.item(0).merge(_tables.item(0).rows.item(j).cells.item(1));
           _cell=_rows.item(j).cells;
           for(k=0;k<_cell.length;k++)
           {
                        if(k==0)
                        {                            
                            app.findGrepPreferences=NothingEnum.nothing;
                            app.changeGrepPreferences=NothingEnum.nothing;
                            app.findGrepPreferences.findWhat="  +";
                            app.changeGrepPreferences.changeTo=" ";
                            app.changeGrep();                          
                            app.findGrepPreferences=NothingEnum.nothing;
                            app.changeGrepPreferences=NothingEnum.nothing;  
                            if(_cell.item(k).contents=="Outpatient surgery \n(Ambulatory Surgical Center/ Hospital)")
                            {
                                _cell.item(k).contents="Outpatient surgery (Ambulatory Surgical Center/Hospital)";
                            }
                            else if(_cell.item(k).contents=="Pediatric eye exam \n(1 visit per year)")
                            {
                                _cell.item(k).contents="Pediatric eye exam (1 visit per year)";   
                            }                            
                            else if(_cell.item(k).contents=="Dental check-up/preventive dental care \n(2 visits per year)")
                            {
                                _cell.item(k).contents="Dental check-up/preventive dental care (2 visits per year)";   
                            }                            
                        
                             _cell.item(k).paragraphs.everyItem().appliedParagraphStyle="Row header";
                         }
                        else if(k==1)
                        {
                            ;

                        }
                        else
                        {
                                app.findTextPreferences=NothingEnum.nothing;
                                app.changeTextPreferences=NothingEnum.nothing;
                                app.findChangeTextOptions.caseSensitive=true;
                                app.findChangeTextOptions.includeMasterPages=true;
                                app.findChangeTextOptions.wholeWord=true;
                                app.findTextPreferences.findWhat="deductible";
                                app.changeTextPreferences.changeTo="ded";
                                app.changeText();
                                app.findTextPreferences=NothingEnum.nothing;
                                app.changeTextPreferences=NothingEnum.nothing;  
                                _cell.item(k).paragraphs.everyItem().appliedParagraphStyle="Table body cell";
                         }
            }
      }

             _tables.item(0).rows.item(17).cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="Row header";
             _tables.item(0).rows.item(17).cells.everyItem().appliedCellStyle="Bold Row";
             _tables.item(0).rows.item(19).cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="Row header";
             _tables.item(0).rows.item(19).cells.everyItem().appliedCellStyle="Bold Row";  
             _tables.item(0).rows.item(22).cells.everyItem().paragraphs.everyItem().appliedParagraphStyle="Row header";
             _tables.item(0).rows.item(22).cells.everyItem().appliedCellStyle="Bold Row";
             _tables.item(0).columns.item(2).cells.everyItem().appliedCellStyle="Column Gap";
             _tables.item(0).columns.item(5).cells.everyItem().appliedCellStyle="Column Gap";  
             _tables.item(0).columns.item(8).cells.everyItem().appliedCellStyle="Column Gap";
     }
             i=i+1;
}
