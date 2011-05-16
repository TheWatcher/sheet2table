var lastelem = null;
var lastinc  = 0;

/** Toggle the header status on a row of cells in the table.
 *  This will determine whether any of the cells on the specified row are already
 *  header cells. If they are, this will remove header status from any cells in
 *  the row that have it. If none of the cells in the row are header cells, this
 *  will set all of the cells to be header cells.
 * 
 *  @param rownum The number of the row to toggle, must be minrow <= rownum <= maxrow.
 */
function toggleHeadRow(rownum)
{
    var hashead = false;

    // Check each element in the row to see whether it has a header set
    // Stop as soon as an element with the header class set is encountered...
    for(var col = mincol; col <= maxcol && !hashead; ++col) {
        var td = $("r"+rownum+"c"+col);
        if(td) hashead = td.hasClass('ishead');
    }

    // Now go through and either remove header from items that have it,
    // or add it to those that don't
    for(var col = mincol; col <= maxcol; ++col) {
        var td = $("r"+rownum+"c"+col);
        if(td) {
            // If we already have a header, we're turning them off
            if(hashead) {
                if(td.hasClass('ishead')) td.removeClass('ishead');

            // If we have no headers, we're turning them on
            } else {
                if(!td.hasClass('ishead')) td.addClass('ishead');
            }
        }
    }
}


/** Compress the header specification in the table into a single string. This
 *  goes through the table, and whenever it encounters a header cell, its position 
 *  is appended to the value of the hidden input element 'hlist', forming a list
 *  of header cell specifications to be sent back to the conversion script.
 * 
 *  @return Always returns true.
 */
function packHeaders() 
{
    var hlist = $('hlist');
    if(!hlist) alert("Missing header list element!");

    hlist.value = ''; // make sure the list is empty before we begin.
    for(var row = minrow; row <= maxrow; ++row) {
        for(var col = mincol; col <= maxcol; ++col) {
            var td = $("r"+row+"c"+col);
            // If we have the element, and is is a header, append it to the list
            if(td && td.hasClass('ishead')) hlist.value += ("r"+row+"c"+col+";");
        }
    }

    return true;
}


/** Set up event handlers on every table data element with the class 'sethead' so
 *  that it behaves appropriately for user-controlled header specification. This 
 *  attaches mouse enter and leave events to modify the apperance of the element,
 *  and a mouse click event that lets the user toggle the header status of the 
 *  element in-situ.
 */
function setHotCells() 
{
    $$('td.sethead').each(function(element, index) {
        element.addEvents({
            mouseover: function() {
                if(!element.hasClass('markpos')) element.addClass('markpos');
            },
            mouseout: function() {
                if(element.hasClass('markpos')) element.removeClass('markpos');
            },
            click: function() {
                if(element.hasClass('ishead')) {
                    element.removeClass('ishead');
                } else {
                    element.addClass('ishead');
                }
                if(lastelem == element) {
                    ++lastinc;
                    if(lastinc == 10) alert("Aren't you getting bored of this yet?");
                    if(lastinc == 15) alert("Seriously, this is getting a bit out of hand, you know...");
                    if(lastinc == 20) alert("What do you have about this cell? Sheesh, all these other cells, but noooo, got to click this one again and again...");
                    if(lastinc == 25) alert("Oh, come on, this isn't even funny now, you're going to wear out this poor cell");
                    if(lastinc == 30) alert("There's a name for people like you, you know?");
                    if(lastinc == 35) alert("Keep this up and you're going to break your mouse button. Or get RSI. And you'll deserve it, you horrible cell molester.");
                    if(lastinc == 40) alert("okay, I'm out of here, this is just too creepy. Weirdo.");
                } else {
                    lastelem = element;
                    lastinc = 0;
                }
            },
        });
    });

    $$('img.toggle').each(function(element, index) {
        var src       = element.getProperty('src');
        var extension = src.substring(src.lastIndexOf('.'),src.length);
        element.addEvent('mouseenter', function() { element.setProperty('src', src.replace(extension, '_over' + extension)); });
        element.addEvent('mouseleave', function() { element.setProperty('src', src); });
    });
}


// Once the dom is available, add the user interaction events to the 
// table data elements.
window.addEvent('domready', function() {
    setHotCells();
});
