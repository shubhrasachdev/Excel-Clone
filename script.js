const PS = new PerfectScrollbar("#cells", {
    wheelSpeed: 2,
    wheelPropagation: true
});

// Getting column and row names/headers
let columns = $("columns");
for(let i = 1; i <= 100; i++){
    let str = "";
    let n = i;
    while(n > 0){
        let rem = n % 26;
        if(rem == 0){ // 0 -> Z, 1 -> A, 2 -> B and so on
            str = 'Z' + str;
            n = Math.floor((n / 26) - 1);
        }else{
            str = String.fromCharCode(rem - 1 + 65) + str; 
            n = Math.floor((n / 26));
        }
    }
    $("#columns").append(`<div class="column-name">${str}</div>`);
    $("#rows").append(`<div class="row-name">${i}</div>`);
}

// Scroll row or column names depending on current scroll on cells
$("#cells").scroll(function(){
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
});

// Generating cells and initializing cell data
let cellData = { "Sheet 1": {} };
let saved = true;
let sheetNumber = 1;
let totalSheets = 1;
let selectedSheet = "Sheet 1";
let defaultProperties = {
    "font-family": "Noto Sans",
    "font-size": 14,
    "text": "",
    "bold": false,
    "italic": false,
    "underlined": false,
    "alignment": "left",
    "color": "#444",
    "bgcolor": "#fff"
};

// For selection by mouse movement
let mousemoved = false, startCellStored = false;
let startCell, endCell;

function addEventsToCells() {

    // Typing allowed only on double click
    $(".input-cell").dblclick(function(){
        unselectCell(this, {}); // unselect all cells other than curr cell
        $(this).attr("contenteditable","true");
        $(this).focus();
    });

    $(".input-cell").blur(function(){
        $(this).attr("contenteditable","false");
        updateCellData("text", $(this).text());
        console.log(cellData);
    });

    // Select / Unselect cells on clicking a cell
    $(".input-cell").click(function (e) {
        if($(this).hasClass("selected")) unselectCell(this, e);
        else selectCell(this, e);
    });

    // selection of cells by mouse move + left click
    mousemoved = false;
    startCellStored = false;
    $(".input-cell").mousemove(function(event) {
        event.preventDefault(); // to prevent cases where selection was failing due to no/little gap b/w previous and current attempt
        if(event.buttons == 1 && !event.ctrlKey) {
            $(".input-cell.selected").removeClass("selected top-selected bottom-selected right-selected left-selected");
            mousemoved = true;
            let [ row, col ] = findRowCol(event.target);
            if(!startCellStored) {
                startCellStored = true;
                startCell = { row: row, col: col };
            } else {
                endCell = { row: row, col: col };
                selectCellsInRange(startCell, endCell);
            }        
        } else if(event.buttons == 0 && mousemoved) {
            startCellStored = false;
            mousemoved = false;
        }
    });
}

function createSheet() {
    $("#cells").empty();
    for(let i = 1; i <= 100; i++){
        let row = $('<div class="cell-row"></div>');
        for(let j = 1; j <= 100; j++){
            row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);
        }
        $('#cells').append(row);
    }
    addEventsToCells();
    addEventsToSheetTabs();
}

createSheet();

function findRowCol(ele) {
    let idArray = $(ele).attr("id").split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);
    return [ rowId, colId ];
}

function getAdjacentCells(ele) {
    let [ rowId, colId ] = findRowCol(ele);
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    let rightCell = $(`#row-${rowId}-col-${colId + 1}`);
    return {topCell, bottomCell, leftCell, rightCell};
}

function selectCell(ele, e, mouseSelection) {
    let {topCell, bottomCell, leftCell, rightCell} = getAdjacentCells(ele);
    if(e.ctrlKey || mouseSelection) {
        // top selected or not
        if(topCell && topCell.hasClass("selected")) {
            topCell.addClass("bottom-selected");
            $(ele).addClass("top-selected");        
        }
        
        // bottom selected or not
        if(bottomCell && bottomCell.hasClass("selected")) {
            bottomCell.addClass("top-selected");
            $(ele).addClass("bottom-selected");
        }

        // left selected or not    
        if(leftCell && leftCell.hasClass("selected")){ 
            leftCell.addClass("right-selected");
            $(ele).addClass("left-selected");
        }
        
        // right selected or not
        if(rightCell && rightCell.hasClass("selected")) {
            rightCell.addClass("left-selected");
            $(ele).addClass("right-selected");
        }
    }  
    else $(".input-cell.selected").removeClass("selected top-selected bottom-selected right-selected left-selected");
    $(ele).addClass("selected");
    changeHeader(findRowCol(ele));
}

// TODO Fix previous attributes of cells being ignored when a new one is selected
function changeHeader([row, col]) {
    let data;
    if(cellData[selectedSheet][row - 1] && cellData[selectedSheet][row - 1][col - 1]) {
        data = cellData[selectedSheet][row - 1][col - 1];
    } else {
        data = defaultProperties;
    }
    $("#font-family").val(data["font-family"]);
    $("#font-family").css("font-family", data["font-family"]);
    $("#font-size").val(data["font-size"]);
    $(".alignment.selected").removeClass("selected");
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");
    addRemoveSelectFromFontStyle(data, "bold");
    addRemoveSelectFromFontStyle(data, "italic");
    addRemoveSelectFromFontStyle(data, "underlined");
    $("#fill-color-icon").css("border-bottom", `3px solid ${data.bgcolor}`);
    $("#text-color-icon").css("border-bottom", `3px solid ${data.color}`);
}

function addRemoveSelectFromFontStyle(data, property) {
    if(data[property]) $(`#${property}`).addClass("selected");
    else $(`#${property}`).removeClass("selected");
}

function unselectCell(ele, e) {
    if(e.ctrlKey && $(ele).attr("contenteditable") == "false") { // unselect only current cell
        let {topCell, bottomCell, leftCell, rightCell} = getAdjacentCells(ele);
        if($(ele).hasClass("top-selected")) topCell.removeClass("bottom-selected");
        if($(ele).hasClass("bottom-selected")) bottomCell.removeClass("top-selected");
        if($(ele).hasClass("left-selected")) leftCell.removeClass("right-selected");
        if($(ele).hasClass("right-selected")) rightCell.removeClass("left-selected");
        $(ele).removeClass("selected top-selected bottom-selected right-selected left-selected");
    } else { // unselect all cells except current cell if CTRL is not pressed.
        $(".input-cell.selected").removeClass("selected top-selected bottom-selected right-selected left-selected");
        $(ele).addClass("selected");
    }
}

function selectCellsInRange(start, end) {
    let sr = Math.min(start.row, end.row);
    let sc = Math.min(start.col, end.col); 
    let er = Math.max(start.row, end.row); 
    let ec = Math.max(start.col, end.col); 
    for(let i = sr; i <= er; i++) {
        for(let j = sc; j <= ec; j++) {
            selectCell($(`#row-${i}-col-${j}`), {}, true)
        }
    }
}

function setFontStyle(ele, property, key, value) {
    if($(ele).hasClass("selected")) {
        $(".input-cell.selected").css(key, "");
        updateCellData(property, false);
        $(ele).removeClass("selected");
    } else {
        $(".input-cell.selected").css(key, value);
        updateCellData(property, true);
        $(ele).addClass("selected");
    }
}

// make cell bold
$("#bold").click(function() {
    setFontStyle(this, "bold", "font-weight", "bold");    
})

// make cell italic
$("#italic").click(function() {
    setFontStyle(this, "italic", "font-style", "italic");    
})

// make cell underlined
$("#underlined").click(function() {
    setFontStyle(this, "underlined", "text-decoration", "underline");    
})

// change alignment of text in cell
$(".alignment").click(function() {
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    let alignment = $(this).attr("data-type");
    $(".input-cell.selected").css("text-align", alignment);

    
});

// change font family or font size
$(".menu-selector").change(function() {
    let value = $(this).val();
    let key = $(this).attr("id");
    if(key == "font-family") $("#font-family").css(key, value);
    if(!isNaN(value)) value = parseInt(value);
    $(".input-cell.selected").css(key, value);
    updateCellData(key, value);
});

// TODO Remove redundancy for font styling here and also above
// color picker for text color and fill color
$(".color-pick").colorPick({
    'initialColor': '#HEX',
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function() {
        if(this.color != "#HEX") {
            if(this.element.attr("id") == "fill-color") {
                $("#fill-color-icon").css("border-bottom", `3.5px solid ${this.color}`);
                $(".input-cell.selected").css("background-color", this.color);
                updateCellData("bgcolor", this.color);

            } else {
                $("#text-color-icon").css("border-bottom", `3.5px solid ${this.color}`);
                $(".input-cell.selected").css("color", this.color);
                updateCellData("color", this.color);

            }
        }
    }
});

function updateCellData(property, value) {
    let prevCellData = JSON.stringify(cellData);
    if(value !== defaultProperties[property]) { // something diff from default value of property has been selected
        $(".input-cell.selected").each(function(_index, data) {
            let [ row, col ] = findRowCol(data);
            if(cellData[selectedSheet][row - 1] == undefined) {
                cellData[selectedSheet][row - 1] = {};
                cellData[selectedSheet][row - 1][col - 1] = {...defaultProperties};
            } else if(cellData[selectedSheet][row - 1][col - 1] == undefined) 
                cellData[selectedSheet][row - 1][col - 1] = {...defaultProperties};
            
            cellData[selectedSheet][row - 1][col - 1][property] = value;
        });
    } else { // default value of property has been selected again
        $(".input-cell.selected").each(function(_index, data) {
            let [ row, col ] = findRowCol(data);
            if(cellData[selectedSheet][row - 1] && cellData[selectedSheet][row - 1][col - 1]) {
                cellData[selectedSheet][row - 1][col - 1][property] = value;
                if(JSON.stringify(cellData[selectedSheet][row - 1][col - 1]) == JSON.stringify(defaultProperties)){
                    delete cellData[selectedSheet][row - 1][col - 1];
                    if(Object.keys(cellData[selectedSheet][row - 1]).length == 0) 
                        delete cellData[selectedSheet][row - 1];
                }
            }
        });
    }
    if(saved && prevCellData != JSON.stringify(cellData)) saved = false;
}

$("#fill-color-icon, #text-color-icon").click(function() {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
});

function selectSheet(ele) {
    $(".sheet-tab.selected").removeClass("selected");
    $(ele).addClass("selected");
    emptySheet();
    selectedSheet = $(ele).text();
    loadSheet();
}

function emptySheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for(let i of rowKeys) {
        let row = parseInt(i);
        let colKeys = Object.keys(data[i]);
        for(let j of colKeys) {
            let col = parseInt(j);
            let cell = $(`#row-${ row + 1 }-col-${ col + 1 }`); // first cell that has changes
            cell.text("");
            cell.css({
                "font-family": defaultProperties["font-family"],
                "font-size": defaultProperties["font-size"] + "px",
                "background-color": defaultProperties["bgcolor"],
                "color": defaultProperties["color"],
                "font-weight": defaultProperties["bold"] ? "bold" : "",
                "font-style": defaultProperties["italic"] ? "italic" : "",
                "text-decoration": defaultProperties["underlined"] ? "underline" : "",
                "text-align": defaultProperties["alignment"]
            });
        } 
    }
}

function loadSheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for(let i of rowKeys) {
        let row = parseInt(i);
        let colKeys = Object.keys(data[i]);
        for(let j of colKeys) {
            let col = parseInt(j);
            let cell = $(`#row-${ row + 1 }-col-${ col + 1 }`); // first cell that has changes
            cell.text(data[row][col].text);
            cell.css({
                "font-family": data[row][col]["font-family"],
                "font-size": data[row][col]["font-size"] + "px",
                "background-color": data[row][col]["bgcolor"],
                "color": data[row][col]["color"],
                "font-weight": data[row][col]["bold"] ? "bold" : "",
                "font-style": data[row][col]["italic"] ? "italic" : "",
                "text-decoration": data[row][col]["underlined"] ? "underline" : "",
                "text-align": data[row][col]["alignment"]
            });
        } 
    }
}

$(".container").click(function() {
    $(".sheet-options-modal").remove();
});

function renameSheet() {
    let newSheetName = $(".sheet-modal-input").val();
    let sheets = Object.keys(cellData);
    if(newSheetName && !sheets.includes(newSheetName)) {
        let newCellData = {};
        for(let sheet of sheets) {
            let key = sheet;
            if(sheet == selectedSheet) key = newSheetName; // sheet for which key has to be renamed
            newCellData[key] = cellData[sheet];
        }
        cellData = newCellData;
        selectedSheet = newSheetName;
        $(".sheet-tab.selected").text(newSheetName);
        $(".sheet-modal-parent").remove();
        saved = false;                    
    } else {
        $(".error").remove();
        $(".sheet-modal-input-container").append(`
            <div class="error"> Sheet name invalid or already exists. </div>
        `);
    }
}

function deleteSheet() {
    if(totalSheets > 1) {
        $(".sheet-modal-parent").remove();
        let sheets = Object.keys(cellData);
        let selectedIdx = sheets.indexOf(selectedSheet);
        let currentSelectedSheet = $(".sheet-tab.selected");
        if(selectedIdx == 0) selectSheet(currentSelectedSheet.next()[0]);
        else selectSheet(currentSelectedSheet.prev()[0]);
        delete cellData[currentSelectedSheet];
        currentSelectedSheet.remove();
        totalSheets--;
        saved = false;
    } else {
        $(".sheet-delete-modal").append(`
            <div class="error"> Sorry, only one sheet exists and cannot be deleted. </div>
        `);
    }
}

function addEventsToSheetTabs() {
    // To give custom context menu
    $(".sheet-tab.selected").bind("contextmenu", function(e) {
        e.preventDefault();
        
        $(".sheet-options-modal").remove();
        let modal = $(`<div class="sheet-options-modal">
                            <div class="option sheet-rename">Rename</div>
                            <div class="option sheet-delete">Delete</div>
                        </div>`);
        $(".container").append(modal);
        $(".sheet-options-modal").css({ "bottom": 0.04 * $(window).height(), "left": e.pageX });
        // can be only added when modal appears, outside will load on launch itself.
    
        $(".sheet-rename").click(function() {
            let renameModal = ` <div class="sheet-modal-parent">
                            <div class="sheet-rename-modal">
                                <div class="sheet-modal-title">
                                    <span> Rename Sheet </span>
                                </div>
                                <div class="sheet-modal-input-container">
                                    <span class="sheet-modal-input-title"> Rename Sheet to: </span>
                                    <input class="sheet-modal-input" type="text" />
                                </div>
                                <div class="sheet-modal-confirmation">
                                    <div class="button ok-button">OK</div>
                                    <div class="button cancel-button">Cancel</div>
                                </div>
                            </div>
                        </div>`;
            $(".container").append(renameModal);
            $(".cancel-button").click(function() {
                $(".sheet-modal-parent").remove();
            });
            $(".ok-button").click(function() {
                renameSheet();
            }); 
            $(".sheet-modal-input").keypress(function(e) {
                if(e.key == "Enter") renameSheet();
            })
        });
        

        $(".sheet-delete").click(function() {
            let deleteModal = `<div class="sheet-modal-parent">
                                <div class="sheet-delete-modal">
                                    <div class="sheet-modal-title">
                                        <span> Delete ${selectedSheet} </span>
                                    </div>
                                    <div class="sheet-modal-detail-container">
                                        <div class="material-icons delete-icon">delete</div>
                                        <span class="sheet-modal-detail-title"> Are you sure you wish to proceed deleting this sheet? </span>
                                    </div>
                                    <div class="sheet-modal-confirmation">
                                        <div class="button delete-button">Delete</div>
                                        <div class="button cancel-button">Cancel</div>
                                    </div>
                                </div>
                            </div>`;
            $(".container").append(deleteModal);
            $(".cancel-button").click(function() {
                $(".sheet-modal-parent").remove();
            });
            $(".delete-button").click(function() {
                deleteSheet();
            }); 
        });

        if(!$(this).hasClass("selected")) selectSheet(this);
    });

    // select a different sheet
    $(".sheet-tab.selected").click(function()  {
        if(!$(this).hasClass("selected")) {
            selectSheet(this);
            $("#row-1-col-1").click();
        }
    });
}

addEventsToSheetTabs();

// Add new sheet
$(".add-sheet").click(function() {
    emptySheet();
    totalSheets++;
    sheetNumber++;
    let sheets = Object.keys(cellData);
    while(sheets.includes("Sheet " + sheetNumber)) sheetNumber++;
    selectedSheet = `Sheet ${sheetNumber}`;
    cellData[selectedSheet] = {};
    $(".sheet-tab.selected").removeClass("selected");
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet ${sheetNumber}</div>`);
    $(".sheet-tab.selected")[0].scrollIntoView();
    addEventsToSheetTabs();
    $("#row-1-col-1").click();
    saved = false;
});

$(".left-scroller").click(function() {
    let sheets = Object.keys(cellData);
    let selectedIdx = sheets.indexOf(selectedSheet);
    if(selectedIdx != 0) {
        selectSheet($(".sheet-tab.selected").prev()[0]);
        selectedSheet = sheets[selectedIdx - 1];
    }
    $(".sheet-tab.selected")[0].scrollIntoView();
});

$(".right-scroller").click(function() {
    let sheets = Object.keys(cellData);
    let selectedIdx = sheets.indexOf(selectedSheet);
    if(selectedIdx != totalSheets - 1) {
        selectSheet($(".sheet-tab.selected").next()[0]);
        selectedSheet = sheets[selectedIdx + 1];
    }
    $(".sheet-tab.selected")[0].scrollIntoView();
});

$("#menu-file").click(function() {
    let fileModal = $(`<div class="file-modal">
        <div class="file-options-modal">
            <div class="close">
                <div class="material-icons file-option-icon close-icon">arrow_circle_down</div>
                <div>Close</div>
            </div>
            <div class="new">
                <div class="material-icons file-option-icon new-icon">insert_drive_file</div>
                <div>New</div>
            </div>
            <div class="open">
                <div class="material-icons file-option-icon open-icon">folder_open</div>
                <div>Open</div>
            </div>
            <div class="save">
                <div class="material-icons file-option-icon save-icon">save</div>
                <div>Save</div>
            </div>
        </div>
        <div class="file-recent-modal"></div>
        <div class="file-transparent-modal"></div>
    </div>`);
    $(".container").append(fileModal);
    fileModal.animate({
        width: "100vw" 
    }, 300);
    $(".close,.new,.save,.open,.file-transparent-modal").click(function() {
        fileModal.animate({
            width: "0vw" 
        }, 300);
        setTimeout(() => {
            fileModal.remove();
        }, 299);
    });

    $(".new").click(function() {
        if(saved) {
            newFile();
        } else {
            let saveModal = `<div class="file-modal-parent">
                                <div class="file-save-modal">
                                    <div class="file-modal-title">
                                        <span> Save ${$(".title-bar").text()} </span>
                                    </div>
                                    <div class="file-modal-detail-container">
                                        You seem to have unsaved changes in the current file. <br/>Do you wish to save them before proceeding?
                                    </div>
                                    <div class="file-modal-confirmation">
                                        <div class="button save-button">Save</div>
                                        <div class="button cancel-button">Cancel</div>
                                    </div>
                                </div>
                            </div>`;
            $(".container").append(saveModal);
            $(".save-button").click(function() {
                $(".file-modal-parent").remove();
                saveFile(true);
            });
            $(".cancel-button").click(function() {
                $(".file-modal-parent").remove();
                newFile();
            });

        }
    })

    $(".save").click(function() {
        saveFile();
    });

    $('.open').click(function() {
        openFile();
    });

});

function newFile() {
    emptySheet();
    $(".sheet-tab").remove();
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet 1 </div>`);
    cellData = {"Sheet 1": {}};
    selectedSheet = "Sheet 1";
    totalSheets = 1;
    sheetNumber = 1;
    addEventsToSheetTabs();
    $("#row-1-col-1").click();
}

function saveFile(createNewFile) {
    $(".container").append(` <div class="sheet-modal-parent">
                                <div class="sheet-rename-modal">
                                    <div class="sheet-modal-title">
                                        <span> Save File </span>
                                    </div>
                                    <div class="sheet-modal-input-container">
                                        <span class="sheet-modal-input-title"> File Name: </span>
                                        <input class="sheet-modal-input" value="${$(".title-bar").text()}" type="text" />
                                    </div>
                                    <div class="sheet-modal-confirmation">
                                        <div class="button ok-button">OK</div>
                                        <div class="button cancel-button">Cancel</div>
                                    </div>
                                </div>
                            </div>`);
    $(".ok-button").click(function() {
        let fileName = $(".sheet-modal-input").val();
        if(fileName) {
            let href = `data:application/json,${encodeURIComponent(JSON.stringify(cellData))}`;
            let a = $(`<a href=${href} download=${fileName + ".json"}/>`);
            $(".container").append(a);
            a[0].click();
            a.remove();
            $(".sheet-modal-parent").remove();
            saved = true;
            if(createNewFile) newFile();
        }
    });

    $(".cancel-button").click(function() {
        $(".sheet-modal-parent").remove();
        if(createNewFile) newFile();
    });
}

function openFile() {
    let inputFile = $(`<input accept="application/json" type="file"/>`);
    $(".container").append(inputFile);
    inputFile.click();
    inputFile.change(function(e) {
        let file = e.target.files[0];
        $(".title-bar").text(file.name.split(".json")[0]); // to change file name to opened file
        let reader = new FileReader();
        reader.readAsText(file);
        reader.onload = function() {
            emptySheet();
            $(".sheet-tab").remove();
            cellData = JSON.parse(reader.result);
            let sheets = Object.keys(cellData);
            for(let sheet of sheets) 
                $(".sheet-tab-container").append(`<div class="sheet-tab selected">${sheet}</div>`);
            addEventsToSheetTabs();
            $(".sheet-tab").removeClass("selected");
            $($(".sheet-tab")[0]).addClass("selected");
            selectedSheet = sheets[0];
            totalSheets = sheets.length;
            sheetNumber = totalSheets;
            loadSheet();
            inputFile.remove();
        }
    });
   
}

let clipboard = {startCell: [], cellData: {}};
let isCut = false;
function copyCells() {
    let selectedCells = $(".input-cell.selected");
    clipboard.startCell = findRowCol(selectedCells[0]);
    clipboard.startCell[0]--;
    clipboard.startCell[1]--;  
    selectedCells.each((_idx, data) => {
        let [row, col] = findRowCol(data);
        if(cellData[selectedSheet][row - 1] && cellData[selectedSheet][row - 1][col - 1]) {
            if(!clipboard.cellData[row - 1]) clipboard.cellData[row - 1] = {};
            clipboard.cellData[row - 1][col - 1] = {...cellData[selectedSheet][row - 1][col - 1]};
        }
    });
    console.log(clipboard);
}

function pasteCells() {
    let startCell = findRowCol($(".input-cell.selected")[0]);
    startCell[0]--;
    startCell[1]--;
    let rows = Object.keys(clipboard.cellData);
    emptySheet();
    for(let row of rows) {
        let cols = Object.keys(clipboard.cellData[row]);
        for(let col of cols) {
            let dx = parseInt(row) - parseInt(clipboard.startCell[0]);
            let dy = parseInt(col) - parseInt(clipboard.startCell[1]);
            let newRow = startCell[0] + dx, newCol = startCell[1] + dy;
            if(!cellData[selectedSheet][newRow]) cellData[selectedSheet][newRow] = {};
            cellData[selectedSheet][newRow][newCol] = {...clipboard.cellData[row][col]};
            if(isCut) {
                delete cellData[selectedSheet][row][col];
                if(Object.keys(cellData[selectedSheet][row]).length == 0) 
                    delete cellData[selectedSheet][row];
            }
        }
    }
    isCut = false;
    loadSheet();
}

$("#cut").click(function() {
    copyCells();
    isCut = true;
});


$("#copy").click(function() {
    copyCells();
});

$("#paste").click(function() {
    pasteCells();
});