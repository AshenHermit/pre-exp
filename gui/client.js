var options = {}
var input_file

window.addEventListener("load", setupPage)

function setupPage(){
    options.table = document.getElementById("generate_table")
    options.graph = document.getElementById("generate_graph")
    input_file = document.getElementById("input_file")
}

function generate(el, type){
    options[type] = {}
    options[type].el = el.parentNode
    eel.generate(type, input_file.value)
}

function open_file(filpath){
    eel.open_file_with_default_program(filpath)
}

eel.expose(on_generated);
function on_generated(type, output_path) {
    if (options[type]!=null){
        options[type].el.querySelector("#result").innerHTML=`
            <div>готово</div>
            <div class="secondary">${output_path}</div>
            <button class="secondary" onclick="open_file('${output_path}')">открыть</button>`
    }
}