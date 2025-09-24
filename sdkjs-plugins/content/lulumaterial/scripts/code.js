(function(window, undefined) {
    window.Asc.plugin.init = function() {
        var parts_dict = {};
        var status="general";
        var parts_left = 0;
        document.getElementById("buttonAddMaterial").onclick = function() {
            console.log("button clicked");
            if(status=="general"){
                var inp_elements = document.getElementsByTagName("input");
                for (let inp in inp_elements){
                    if (inp_elements[inp].checked){
                        parts_dict[inp_elements[inp].name]=[];
                    }
                }
                
                parts_left=Object.keys(parts_dict).length
                if (parts_left>0){
                    console.log(parts_left)
                    var part=Object.keys(parts_dict)[0]
                    $(function(){$("#includedContent").load("form-configuration.html");});
                    if (part=="столешница"){
                        $(function(){$("#field_material").load("select-material-tree-stol.html");});
                    }
                    document.getElementById("p_legend").innerHTML="Конфигурация элемента \""+ part + "\"";
                    var div = document.createElement("div");
                    div.setAttribute("name", part);
                    div.setAttribute("id", "appendedDiv");
                    document.getElementById("myForm").appendChild(div);
                    
                    var rad = document.getElementsByName("surface");
                    rad[0].onclick = function() {
                        if (part=="столешница"){
                            $(function(){$("#field_material").load("select-material-tree-stol.html");});
                        }
                        else{
                            $(function(){$("#field_material").load("select-material-tree-not-stol.html");});
                        }
                    };
                    rad[1].onclick = function() {
                        $(function(){$("#field_material").load("select-material-smooth.html");});
                    };
                    rad[2].onclick = function() {
                        $(function(){$("#field_material").load("select-material-other.html");});
                    };
                    var custom_color_inp = document.getElementById("ral-ncs-code");
                    custom_color_inp.onclick = function() {
                        document.getElementById("custom-code").checked = true;
                    };
                    status="surface";
                }
                else {
                    Asc.scope.parts=''
                    for (let part in parts_dict){
                        Asc.scope.parts =  Asc.scope.parts + part + " ";
                    }
        
                    window.Asc.plugin.callCommand(function() {
                        var oWorksheet = Api.GetActiveSheet();
                        var ActiveCell = oWorksheet.ActiveCell;
                        ActiveCell.SetValue(Asc.scope.parts);
                    }, true);
                }
            }
            

        };
    };
    
    window.Asc.plugin.button = function(id) {
        console.log(id);
        if (id==-1){
            this.executeCommand("close", "");
        }
    };

})(window, undefined);
