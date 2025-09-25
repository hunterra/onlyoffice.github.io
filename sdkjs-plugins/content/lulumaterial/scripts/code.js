(function(window, undefined) {
    window.Asc.plugin.init = function() {
        var parts_dict = {};
        var status="general";
        var parts_left = 0;
        var part_name = '';
        
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
                Asc.scope.parts_dict={};
                var inp_elements = document.getElementsByTagName("input");
                for (let inp in inp_elements){
                    if (inp_elements[inp].checked && !exclude_array.includes(inp_elements[inp].name)){
                        parts_dict[inp_elements[inp].name]=[];
                    }
                }
                parts_left=Object.keys(parts_dict).length;
                console.log(parts_dict);
            }
            if(status=="surface"){
                var selected_radio_surface = document.querySelector('input[name="surface"]:checked').nextElementSibling.innerText;
                var selected_radio_material = document.querySelector('input[name="material"]:checked').nextElementSibling.innerText;
                var selected_radio_color = document.querySelector('input[name="color"]:checked')
                if (selected_radio_color.id=="custom-code"){
                    selected_radio_color = "RALNCSCODE" + document.getElementById("ral-ncs-code").value
                }
                else {
                    selected_radio_color = selected_radio_color.nextElementSibling.innerText;
                }
                
                if(selected_radio_material=="Плита МДФ в обкладке из массива дуба, покрытая шпоном дуба"){
                    selected_radio_material= "Плита МДФ в обкладке из массива дуба, покрытая шпоном дуба (шпон 1,5 мм)"
                }
                
                Asc.scope.parts_dict[part_name]=[selected_radio_surface,selected_radio_material,selected_radio_color];
                parts_left=parts_left-1;
            }
            if(parts_left>0){
                var part=Object.keys(parts_dict)[0];
                document.getElementById("p_legend").innerHTML="Конфигурация элемента  \""+part +"\"";
                partSelectionForm.style.display = 'none';
                partConfigurationForm.style.display = 'block';
                part_name = part;
                delete parts_dict[part];
                    
                status="surface";
            }
            else{
                console.log(Asc.scope.parts_dict);
                window.Asc.plugin.callCommand(function() {
                var oWorksheet = Api.GetActiveSheet();
                var ActiveCell = oWorksheet.ActiveCell;
                ActiveCell.SetValue(Asc.scope.parts_dict);
                }, true);
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
