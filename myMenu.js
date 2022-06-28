var myColors = [
    "#bfbfbf", "#a6a6a6", "#7f7f7f", "#E223EB", "#b973a6", "#a385e2",
    "#8085e9", "#76ba99", "#4bacc6", "#00A5FF", "#ff0000", "#F76B43",
    "#d7767f", "#f15c80"
];

function addColorMenu() {
    var span = document.createElement("span");
    span.setAttribute("class", "menu-item-container");
    span.setAttribute("style", "margin-top: 10px");

    var divColor;
    for (var j = 0; j < myColors.length; j++) {
        if (j % 6 == 0) {
            divColor = document.createElement("div");
            span.appendChild(divColor);
        }

        var color = myColors[j];
        var id = "myColor_" + color.replace("#", "");

        var div = document.createElement("div");
        div.setAttribute("style", "background-color:" + color);
        div.setAttribute("class", "color_div");

        var a = document.createElement("a");
        a.setAttribute("onclick", "javascript:setColor('" + id + "', '" + color + "');");
        a.appendChild(div);

        var aHide = document.createElement("a");
        aHide.setAttribute("class", "menu-style-btn");
        aHide.setAttribute("data-style", id);
        aHide.setAttribute("data-label", "Color[" + color + "]");
        aHide.setAttribute("data-lg", "Menu");
        aHide.setAttribute("style", "display:none");
        aHide.setAttribute("id", id);

        divColor.appendChild(a);
        divColor.appendChild(aHide);
    }

    var style = document.createElement("style");
    style.innerHTML = ".color_div{width:25px; height:20px; margin-right:5px; display:inline-block;}";

    var liColor = document.createElement("li");
    liColor.setAttribute("data-group", "inline-style");
    liColor.setAttribute("class", "has-btn-submenu");
    liColor.appendChild(style);
    liColor.appendChild(span);

    var liLine = document.createElement("li");
    liLine.setAttribute("class", "divider");

    var contextMenu = document.getElementById("context-menu");
    contextMenu.innerHTML = liColor.outerHTML + liLine.outerHTML + contextMenu.innerHTML;
}

function setColor(id, color) {

    File.editor.stylize.CONST.SyntaxMap[id] = ["<font color='" + color + "'>", "</font>"];
    File.editor.stylize.CONST.StyleToNameMap[id] = "Color[" + color + "]";

    var a = document.getElementById(id);
    a.click();
}

(function(){
    addColorMenu();
})();