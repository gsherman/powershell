
function header1([string]$str){
 "<h1>" + $str + "</h1>"
}

function header2([string]$str){
 "<h2>" + $str + "</h2>"
}


function header3([string]$str){
 "<h3>" + $str + "</h3>"
}

function header4([string]$str){
 "<h4>" + $str + "</h4>"
}

function paragraph([string]$str){
 "<p>" + $str + "</p>"
}

function div([string]$str, [string]$cls){
 "<div class=" + $cls + ">" + $str
}

function endDiv(){
 "</div>"
}

function htmlBreak(){
 "<br/>"
}

function footer([string]$str){
 "<div class=footer>" + $str + "</div>"
}
