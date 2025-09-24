(function(window, undefined) {
    window.Asc.plugin.init = function() {
        var parts_dict = {};
        var status="general";
        var parts_left = 0;
        
        var partSelectionForm = document.getElementById("partSelection");
        var partConfigurationForm = document.getElementById("partConfiguration");

        var mdfDiv = document.getElementById("div-mdf");
        var massivDiv = document.getElementById("div-massiv");
        var mdfDubDiv = document.getElementById("div-mdf-dub");
        var dopStolMaterialDiv = document.getElementById("div-dop-stol-material");
        
        var massivBukaDiv = document.getElementById("div-massiv-buka");
        var mdfEmalDiv = document.getElementById("div-mdf-emal");
        var mdfLamDiv = document.getElementById("div-mdf-lam");
        
        var porolonDiv = document.getElementById("div-porolon");
        var metalDiv = document.getElementById("div-metal");
        
        var exclude_array = ["surface", "material", "color"];
        
        document.getElementById("buttonAddMaterial").onclick = function() {
            console.log("button clicked");
            if(status=="general"){
                var inp_elements = document.getElementsByTagName("input");
                for (let inp in inp_elements){
                    if (inp_elements[inp].checked && !exclude_array.includes(inp_elements[inp].name)){
                        parts_dict[inp_elements[inp].name]=[];
                    }
                }
                parts_left=Object.keys(parts_dict).length;
                console.log(parts_left);
            }
            if(parts_left>0){
                let part=parts_dict[0];
                        console.log(part)
                        document.getElementById("p_legend").innerHTML="Конфигурация элемента  \""+part +"\"";
                        partSelectionForm.style.display = 'none';
                        partConfigurationForm.style.display = 'block';
                        parts_left=parts_left-1;
                        delete parts_dict.part;
                    
                    
                status="surface";
            }
            else{
                console.log("no parts left")
            }

            }
            

        };

    
    window.Asc.plugin.button = function(id) {
        console.log(id);
        if (id==-1){
            this.executeCommand("close", "");
        }
    };

})(window, undefined);
