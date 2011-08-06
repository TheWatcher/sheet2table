var imgpath  = "templates/default/images/icons/";  //!< The location of the header images 
var pickmode = "anchor";  //!< current mouse click mode, will be either "anchor" or "body"
var startcol = -1;        //!< When in "body" mode, this is the column of the matching anchor
var modes    = new Array; //!< each element contains the current mode string for the column it corresponds to
var anchors  = new Array; //!< each element contains the popup id of the anchor, or -1 
var bodies   = new Array; //!< each element contains the popup id of the body, or -1

/** Add mouse-over event handlers to the header images.
 *  This will go through each of the popup header images and attach
 *  mouse enter and mouse leave events that will change the images
 *  to indicate clickability.
 */
function setHotCells() 
{
    $$('img.pophead').each(function(element, index) {
	    element.addEvent('mouseenter', function() { 
            var id  = element.getProperty('id');
            var col = id.substring(6).toInt();
           
            element.setProperty('src', imgpath + modes[col] + "_over.png");
        });
	    element.addEvent('mouseleave', function() { 
            var id  = element.getProperty('id');
            var col = id.substring(6).toInt();
           
            element.setProperty('src', imgpath + modes[col] + ".png");
        });
    });
}

// When the dom is ready, add in the appropriate event handlers.
window.addEvent('domready', function() {
    setHotCells();
});


/** Change all popup headers in a specified mode to another.
 *  This will go through all the popup headers and, if they match the 'replace'
 *  mode, they will be set to the 'target' mode instead, provided that the
 *  column does not match the trigger column.
 * 
 *  @param triggercol The number of the column that triggered the change.
 *  @param replace    The mode to search through the headers for.
 *  @param target     The mode to set matching headers to.
 */
function setColHeads(triggercol, replace, target)
{
    for(var col = mincol; col <= maxcol; ++col) {
        // Leave the column that triggered the change alone
        if(col != triggercol) {
            // If the column is in the mode we want to replace, do it
            if(modes[col] == replace) setColMode(col, target, '');
        }
    }
}


/** Update the mode of the specified column. This will set the column to
 *  the specified mode, and update the header image accordingly.
 * 
 *  @param col   The column to update.
 *  @param mode  The mode to set the column to.
 *  @param title The title to set on the column image.
 */
function setColMode(col, mode, title)
{
    modes[col] = mode;
                
    var element = $("pophot"+col);
    if(element) {
        element.setProperty('src', imgpath + modes[col] + ".png");
        element.setProperty('title', title);
    }
}


/** Add the specified class to all rows of cells in a column.
 *  This will go through each cell in the specified column, and if the cell does
 *  not have the provided class set, it will be added to the cell.
 * 
 *  @param col      The column of cells to add the class to.
 *  @param setclass The class to add to cells, if they do not already have it.
 */
function setColRows(col, setclass)
{
    for(var row = minrow; row <= maxrow; ++row) {
        var element = $("r"+row+"c"+col);
        if(element && !element.hasClass(setclass)) element.addClass(setclass);
    }
}


/** Remove the specified class from all rows of cells in a column.
 *  This will go through each cell in the specified column, and if the cell
 *  has the provided class set, it will be removed from the cell.
 * 
 *  @param col      The column of cells to add the class to.
 *  @param remclass The class to remove from cells, if they have it set.
 */
function clearColRows(col, remclass)
{
    for(var row = minrow; row <= maxrow; ++row) {
        var element = $("r"+row+"c"+col);
        if(element && element.hasClass(remclass)) element.removeClass(remclass);
    }
}


/** Clear the anchor and body columns associated with the specified 
 *  popup id. This will search the columns for the specified popup id
 *  and reset the anchor and body columns to default values, destroying the
 *  popup. 
 * 
 *  @param id The id of the popup to remove from the page.
 */
function clearPopup(id)
{
    var acol = -1;
    var bcol = -1;

    // First determien which columns the specified popup id has as its anchor and body
    for(var col = mincol; col <= maxcol && (acol == -1 || bcol == -1); ++col) {
        if(anchors[col] == id) acol = col;
        if(bodies[col] == id) bcol = col;
    }

    // If either column is -1, we fail
    if(acol == -1 || bcol == -1) {
        alert("Unable to find anchor or body columns for popup id "+id+". This should not happen.");
        return;
    }

    // clear the column highlight
    clearColRows(acol, "isanchor");
    clearColRows(bcol, "isbody");

    // And reset the popup.
    anchors[acol] = -1;
    bodies[bcol] = -1;
    setColMode(acol, "anchor_add", '');
    setColMode(bcol, "anchor_add", '');

    // Note that we do not mess with nextid here! The popup we nuked may not be 
    // nextid-1, just let the perl script deal with handling the mess.
}


/** Cancel the current in-progress popup creation operation. This will reset the
 *  state of the pick mode and start column variables to defaults, and clear the
 *  partial popup data back to the initial values.
 * 
 *  @param col The column that initiated the popup create operation (the current popup anchor) 
 */
function cancelPopup(col)
{
    pickmode = "anchor";
    startcol = -1;
    setColHeads(col, "body_add", "anchor_add");
    clearColRows(col, "isanchor");
    anchors[col] = -1;
}


/** Handle the user's clicks on popup columns. This will handle the user's clicks on each
 *  popup column and either create popups or remove them depending on the context.
 * 
 *  @param col The column the user has clicked on.
 */
function handlePopClick(col)
{ 
    // If the col is an anchor col, we need to clear its popupid
    if(modes[col] == "anchor")  {
        // If we're in anchor pick mode, we are deleting the popup
        if(pickmode == "anchor") {
            clearPopup(anchors[col]);
        
        // If we're in body pick mode, cancel the popup setup
        } else {
            cancelPopup(startcol);
        }

    } else if(modes[col] == "body") {
        // If we're in anchor pick mode, we are deleting the popup
        if(pickmode == "anchor") {
            clearPopup(bodies[col]);
        
        // If we're in body pick mode, cancel the popup setup
        } else {
            cancelPopup(startcol);
        }
        
    } else if(modes[col] == "anchor_add") {
        // If we are picking an anchor, we need to change the mode, set the
        // selected column as the new id anchor, and change all other 
        // anchor_add columns to body_add columns
        if(pickmode == "anchor") {
            pickmode = "body";
            startcol = col;
            setColHeads(col, "anchor_add", "body_add");
            setColRows(col, "isanchor");
            anchors[col] = nextid;
            
        // If we're in pick body mode, and the user has clicked on the add (there will 
        // only be one - the one that triggered the body add) we need to cancel the add
        } else {
            cancelPopup(startcol);
        }
            
    } else if(modes[col] == "body_add") {
        // If we are in pick mode and a body add has been clicked, the user has selected
        // the body column for a popup, so mark it
        if(pickmode == "body") {
            pickmode = "anchor";
            bodies[col] = nextid;

            // clear the other columns back to anchor add mode if needed
            setColHeads(col, "body_add", "anchor_add");

            // Now set the rows and column headers
            setColRows(col, "isbody");
            setColMode(col, "body", "Popup #"+nextid);
            setColMode(startcol, "anchor", "Popup #"+nextid);

            ++nextid;
            startcol = -1;
        }
    }
}


/** Determine which column corresponds to the body column for the specified
 *  popup id. This will return the column number if found, or -1 if there is
 *  no body for the popup.
 * 
 *  @param id The popup id to search for.
 *  @return The column number of the body for the specified popup id, or -1.
 */
function findBodyCol(id)
{
    for(var col = mincol; col <= maxcol; ++col) {
        if(bodies[col] == id) return col;
    }
    return -1;
}


/** Compress the popup column information into a string to send back to the 
 *  cgi script. This will go through the popups on the page and store their 
 *  anchor and body column ids in a form the cgi script can interpret.
 * 
 *  @return Always returns true. 
 */
function packPopups()
{
    var plist = $('plist');
    if(!plist) alert("Missing popup list element!");

    // cancel pick mode if needed
    if(pickmode == "body" && startcol > -1) cancelPopup(startcol);    

    plist.value = ''; // make sure the list is empty before we begin.
    for(var col = mincol; col <= maxcol; ++col) {
        // If we find an anchor, we want to find the matching body
        if(anchors[col] > -1) {
            var bodycol = findBodyCol(anchors[col]);            

            // Only add the popup if we have both the anchor and body (ie: handle the 
            // situation where the user has hit Next while in body add mode)
            if(bodycol > -1) {
                plist.value += "a"+col+"b"+bodycol+";";
            }
        }
    }

    return true;
}