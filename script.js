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

// Generating cells and initializing cell data
let cellData = [];
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
    cellData.push(rowArray);
    $('#cells').append(row);
}

// Scroll row or column names depending on current scroll on cells
$("#cells").scroll(function(){
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
});

// Typing allowed only on double click
$(".input-cell").dblclick(function(){
    unselectCell(this, {}); // unselect all cells other than curr cell
    $(this).attr("contenteditable","true");
    $(this).focus();
});

$(".input-cell").blur(function(){
    $(this).removeClass("selected");
    $(this).attr("contenteditable","false");
});

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
    let data = cellData[row - 1][col - 1];
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

// Select / Unselect cells on clicking a cell
$(".input-cell").click(function (e) {
    if($(this).hasClass("selected")) unselectCell(this, e);
    else selectCell(this, e);
});

// selection of cells by mouse move + left click
let mousemoved = false, startCellStored = false;
let startCell, endCell;
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
})

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
            cellData[row - 1][col - 1][property] = false;
        })
        $(ele).removeClass("selected");
    } else {
        $(".input-cell.selected").css(key, value);
        $(".input-cell.selected").each(function(_idx, data) {
            let [ row, col ] = findRowCol(data);
            cellData[row - 1][col - 1][property] = true;
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
        cellData[row - 1][col - 1].alignment = alignment;
    });
});

// change font family or font size
$(".menu-selector").change(function(e) {
    let value = $(this).val();
    let key = $(this).attr("id");
    if(key == "font-family") $("#font-family").css(key, value);
    if(!isNaN(value)) value = parseInt(value);
    $(".input-cell.selected").css(key, value);
    $(".input-cell.selected").each(function(_idx, data) {
        let [row, col] = findRowCol(data);
        cellData[row - 1][col - 1][key] = value;
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
                    cellData[row - 1][col - 1].bgcolor = this.color;
                });

            } else {
                $("#text-color-icon").css("border-bottom", `3.5px solid ${this.color}`);
                $(".input-cell.selected").css("color", this.color);
                $(".input-cell.selected").each((_idx, data) => {
                    let [row, col] = findRowCol(data);
                    cellData[row - 1][col - 1].color = this.color;

                });
            }
        }
    }
});

$("#fill-color-icon, #text-color-icon").click(function() {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
})