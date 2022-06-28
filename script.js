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
let cellData = { "Sheet 1": [] };
let sheetNumber = 1;
let totalSheets = 1;
let selectedSheet = "Sheet 1";

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
        $(this).removeClass("selected");
        $(this).attr("contenteditable","false");
        let [row, col] = findRowCol(this);
        cellData[selectedSheet][row - 1][col - 1].text = $(this).text();
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
        let rowArray = [];
        for(let j = 1; j <= 100; j++){
            row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);
            rowArray.push({
                "font-family": "Noto Sans",
                "font-size": 14,
                "text": "",
                "bold": false,
                "italic": false,
                "underlined": false,
                "alignment": "left",
                "color": "#444",
                "bgcolor": "#fff"
            });
        }
        cellData[selectedSheet].push(rowArray);
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
    let data = cellData[selectedSheet][row - 1][col - 1];
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
        $(".input-cell.selected").each(function(_idx, data) {
            let [ row, col ] = findRowCol(data);
            cellData[selectedSheet][row - 1][col - 1][property] = false;
        })
        $(ele).removeClass("selected");
    } else {
        $(".input-cell.selected").css(key, value);
        $(".input-cell.selected").each(function(_idx, data) {
            let [ row, col ] = findRowCol(data);
            cellData[selectedSheet][row - 1][col - 1][property] = true;
        })
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
    $(".input-cell.selected").each(function(_index, data) {
        let [ row, col ] = findRowCol(data);
        cellData[selectedSheet][row - 1][col - 1].alignment = alignment;
    });
});

// change font family or font size
$(".menu-selector").change(function() {
    let value = $(this).val();
    let key = $(this).attr("id");
    if(key == "font-family") $("#font-family").css(key, value);
    if(!isNaN(value)) value = parseInt(value);
    $(".input-cell.selected").css(key, value);
    $(".input-cell.selected").each(function(_idx, data) {
        let [row, col] = findRowCol(data);
        cellData[selectedSheet][row - 1][col - 1][key] = value;
    });
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
                $(".input-cell.selected").each((_idx, data) => {
                    let [row, col] = findRowCol(data);
                    cellData[selectedSheet][row - 1][col - 1].bgcolor = this.color;
                });

            } else {
                $("#text-color-icon").css("border-bottom", `3.5px solid ${this.color}`);
                $(".input-cell.selected").css("color", this.color);
                $(".input-cell.selected").each((_idx, data) => {
                    let [row, col] = findRowCol(data);
                    cellData[selectedSheet][row - 1][col - 1].color = this.color;

                });
            }
        }
    }
});

$("#fill-color-icon, #text-color-icon").click(function() {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
});

function selectSheet(ele) {
    addLoader();
    $(".sheet-tab.selected").removeClass("selected");
    $(ele).addClass("selected");
    selectedSheet = $(ele).text();
    setTimeout(() => {
        loadSheet();
        removeLoader();
    });
}

function loadSheet() {
    $("#cells").empty();
    let data = cellData[selectedSheet];
    for(let i = 1; i <= data.length; i++) {
        let row = $('<div class="cell-row"></div>');
        for(let j = 1; j < data[i - 1].length; j++) {
            let cell = $(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false">${data[i - 1][j - 1]["text"]}</div>`);
            cell.css({
                "font-family": data[i - 1][j - 1]["font-family"],
                "font-size": data[i - 1][j - 1]["font-size"] + "px",
                "background-color": data[i - 1][j - 1]["bgcolor"],
                "color": data[i - 1][j - 1]["color"],
                "font-weight": data[i - 1][j - 1]["bold"] ? "bold" : "",
                "font-style": data[i - 1][j - 1]["italic"] ? "italic" : "",
                "text-decoration": data[i - 1][j - 1]["underlined"] ? "underline" : "",
                "text-align": data[i - 1][j - 1]["alignment"]
            });
            row.append(cell);
        } 
        $("#cells").append(row);
    }
    addEventsToCells();
}

$(".container").click(function() {
    $(".sheet-options-modal").remove();
});

function renameSheet() {
    let value = $(".sheet-modal-input").val();
    if(value && !Object.keys(cellData).includes(value)) {
        cellData[value] = cellData[selectedSheet];
        delete cellData[selectedSheet];
        selectedSheet = value;
        $(".sheet-tab.selected").text(value);
        $(".sheet-modal-parent").remove();                    
    } else {
        $(".error").remove();
        $(".sheet-modal-input-container").append(`
            <div class="error"> Sheet name invalid or already exists. </div>
        `);
    }
}

function deleteSheet() {
    if(totalSheets > 1) {
        let sheets = Object.keys(cellData);
        $(".sheet-modal-parent").remove();
        let selectedIdx = sheets.indexOf(selectedSheet);
        let currentSelectedSheet = $(".sheet-tab.selected");
        delete cellData[selectedSheet];
        if(selectedIdx == 0) selectSheet(currentSelectedSheet.next()[0]);
        else selectSheet(currentSelectedSheet.prev()[0]);
        currentSelectedSheet.remove();
        totalSheets--;
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
                                        <span> Delete Sheet </span>
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
        if(!$(this).hasClass("selected")) selectSheet(this);
    });
}

addEventsToSheetTabs();

// Add new sheet
$(".add-sheet").click(function() {
    addLoader();
    totalSheets++;
    sheetNumber++;
    let sheets = Object.keys(cellData);
    while(sheets.includes("Sheet " + sheetNumber)) sheetNumber++;
    selectedSheet = `Sheet ${sheetNumber}`;
    cellData[selectedSheet] = [];
    $(".sheet-tab.selected").removeClass("selected");
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet ${sheetNumber}</div>`);
    setTimeout(() => {
        createSheet();
        removeLoader();
    }, 10);
    $(".sheet-tab.selected")[0].scrollIntoView();
});

function addLoader() {
  $(".container").append(`
            <div class="loader-parent">
                    <span class="loader"/>
                </div>`);
}

function removeLoader() {
    $(".loader-parent").remove();
}

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




