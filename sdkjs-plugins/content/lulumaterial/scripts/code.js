(function(window, undefined) {
    window.Asc.plugin.init = function() {
        var parts_dict = {};
        var status="general";
        var parts_left = 0;
        var part_name = "";
        
        var partSelectionForm = document.getElementById("partSelection");
        var partConfigurationForm = document.getElementById("partConfiguration");
        
        var backButton = document.getElementById("back_button");

        var mdfDiv = document.getElementById("div-mdf");
        var massivDiv = document.getElementById("div-massiv");
        var mdfDubDiv = document.getElementById("div-mdf-dub");
        var dopStolMaterialDiv = document.getElementById("div-dop-stol-material");
        
        var massivBukaDiv = document.getElementById("div-massiv-buka");
        var mdfEmalDiv = document.getElementById("div-mdf-emal");
        var mdfLamDiv = document.getElementById("div-mdf-lam");
        
        var porolonDiv = document.getElementById("div-porolon");
        var metalDiv = document.getElementById("div-metal");
        
        var mdfInp = document.getElementById("mdf");
        var massivBukaInp = document.getElementById("massiv-buka");
        var porolonInp = document.getElementById("porolon");
        
        var natDubColor = document.getElementById("nat-dub");
        var ralNcsCode = document.getElementById('ral-ncs-code');
        ralNcsCode.onclick = function() {
                document.getElementById("custom-code").checked = true;
            };

        
        var exclude_array = ["surface", "material", "color"];
        
        var rad = document.getElementsByName("surface");
        rad[0].onclick = function() {
            
            mdfDiv.style.display = 'block';
            massivDiv.style.display = 'block';
            mdfDubDiv.style.display = 'block';
            if (part_name=="столешница"){
                dopStolMaterialDiv.style.display = 'block';
            }
            massivBukaDiv.style.display = 'none';
            mdfEmalDiv.style.display = 'none';
            mdfLamDiv.style.display = 'none';
            
            porolonDiv.style.display = 'none';
            metalDiv.style.display = 'none';
            
            mdfInp.checked = true;
        };
        rad[1].onclick = function() {
            mdfDiv.style.display = 'none';
            massivDiv.style.display = 'none';
            mdfDubDiv.style.display = 'none';
            dopStolMaterialDiv.style.display = 'none';
            
            massivBukaDiv.style.display = 'block';
            mdfEmalDiv.style.display = 'block';
            mdfLamDiv.style.display = 'block';
            
            porolonDiv.style.display = 'none';
            metalDiv.style.display = 'none';
            
            massivBukaInp.checked = true;
        };
        rad[2].onclick = function() {
            mdfDiv.style.display = 'none';
            massivDiv.style.display = 'none';
            mdfDubDiv.style.display = 'none';
            dopStolMaterialDiv.style.display = 'none';
            
            massivBukaDiv.style.display = 'none';
            mdfEmalDiv.style.display = 'none';
            mdfLamDiv.style.display = 'none';
            
            porolonDiv.style.display = 'block';
            metalDiv.style.display = 'block';
            
            porolonInp.checked = true;
        };
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
                    selected_radio_color = "RALNCSCODE" + ralNcsCode.value;
                }
                else {
                    selected_radio_color = selected_radio_color.nextElementSibling.innerText;
                }
                
                if(selected_radio_material=="Плита МДФ в обкладке из массива дуба, покрытая шпоном дуба"){
                    selected_radio_material= "Плита МДФ в обкладке из массива дуба, покрытая шпоном дуба (шпон 1,5 мм)";
                }
                
                Asc.scope.parts_dict[part_name]=[selected_radio_surface,selected_radio_material,selected_radio_color];
                parts_left=parts_left-1;
            }
            if(parts_left>0){
                var part=Object.keys(parts_dict)[0];
                part_name = part;
                document.getElementById("p_legend").innerHTML="Конфигурация элемента  \""+part_name +"\"";
                partSelectionForm.style.display = 'none';
                partConfigurationForm.style.display = 'block';
                backButton.style.display = 'block';
                natDubColor.checked = true;
                ralNcsCode.value = '';
                rad[0].click();
                
                delete parts_dict[part];
                    
                status="surface";
            }
            else{
                console.log(Asc.scope.parts_dict);
                Asc.scope.cell_val = "";
                var smooth_dict = {};
                var tree_dict = {};
                var no_surf_dict = {};
                var color_dict = {};
                Object.keys(Asc.scope.parts_dict).forEach(function(value) {
                    var ar = Asc.scope.parts_dict[value];
                    if(ar[0]=="Поверхности гладкие"){
                        if(ar[1] in smooth_dict){
                            smooth_dict[ar[1]].push(value);
                        }
                        else{
                            smooth_dict[ar[1]]=[value];
                        }
                    }
                    else if(ar[0]=="Поверхности со структурой дерева"){
                        if(ar[1] in tree_dict){
                            tree_dict[ar[1]].push(value);
                        }
                        else{
                            tree_dict[ar[1]]=[value];
                        }
                    }
                    else {
                        if(ar[1] in no_surf_dict){
                            no_surf_dict[ar[1]].push(value);
                        }
                        else{
                            no_surf_dict[ar[1]]=[value];
                        }
                    }
                    if(ar[2] in color_dict){
                        color_dict[ar[2]].push(value);
                    }
                    else{
                        color_dict[ar[2]]=[value];
                    }
                });
                
                Asc.scope.cell_val = "";
                if(Object.keys(tree_dict).length>0){
                    Asc.scope.boldCharList = [[0,30]];
                    Asc.scope.cell_val = "ПОВЕРХНОСТИ С ТЕКСТУРОЙ ДЕРЕВА\n\n";
                    Object.keys(tree_dict).forEach(function(value) {
                    var decap_value = value.substring(0,1).toLowerCase() + value.substring(1);
                    Asc.scope.cell_val = Asc.scope.cell_val + tree_dict[value].join(', ') + ' – ' + decap_value + '\n\n';
                    });
                }
                if(Object.keys(smooth_dict).length>0){
                    Asc.scope.boldCharList.push([Asc.scope.cell_val.length,20]);
                    Asc.scope.cell_val = Asc.scope.cell_val + "ПОВЕРХНОСТИ ГЛАДКИЕ\n\n";
                    Object.keys(smooth_dict).forEach(function(value) {
                    var decap_value = value.substring(0,1).toLowerCase() + value.substring(1);
                    var box_index = smooth_dict[value].indexOf('внутренние ящики')
                    if(box_index>=0){
                        smooth_dict[value].splice(box_index,1);
                        smooth_dict[value].push('внутренние ящики');
                    }
                    Asc.scope.cell_val = Asc.scope.cell_val + smooth_dict[value].join(', ') + ' – ' + decap_value + '\n\n';
                    });
                }
                if(Object.keys(no_surf_dict).length>0){
                    Asc.scope.boldCharList.push([Asc.scope.cell_val.length,10]);
                    Asc.scope.cell_val = Asc.scope.cell_val +  "ОСТАЛЬНОЕ\n\n";
                    Object.keys(no_surf_dict).forEach(function(value) {
                    var decap_value = value.substring(0,1).toLowerCase() + value.substring(1);
                    Asc.scope.cell_val = Asc.scope.cell_val + no_surf_dict[value].join(', ') + ' – ' + decap_value + '\n\n';
                    });
                }
                var custom_color=false;
                var temp_value = "";
                Asc.scope.color_cell_val = "";
                Asc.scope.colorBoldCharList = [];
                Object.keys(color_dict).forEach(function(value) {
                    if(value.startsWith("RALNCSCODE")){
                        custom_color=true;
                        temp_value = value.replace("RALNCSCODE","");
                        Asc.scope.color_cell_val = Asc.scope.color_cell_val + color_dict[value].join(', ') + ' – ' + temp_value + '\n\n';
                        Asc.scope.colorBoldCharList.push([Asc.scope.color_cell_val.length - temp_value.length - 1,temp_value.length]);
                    }
                    else {
                        Asc.scope.color_cell_val = Asc.scope.color_cell_val + color_dict[value].join(', ') + ' – ' + value + '\n\n';
                        Asc.scope.colorBoldCharList.push([Asc.scope.color_cell_val.length - value.length - 1,value.length]);
                    }
                    });
                
                
                
                /*
                for (let asc_part in Asc.scope.parts_dict){
                    Asc.scope.cell_val = Asc.scope.cell_val + asc_part + ": ";
                    for (let val in Asc.scope.parts_dict[asc_part]){
                        Asc.scope.cell_val = Asc.scope.cell_val + Asc.scope.parts_dict[asc_part][val] + ", ";
                    }
                    Asc.scope.cell_val = Asc.scope.cell_val.slice(0, -2) + "\n";
                }
                */
                window.Asc.plugin.callCommand(function() {
                    var redColor = Api.CreateColorFromRGB(255, 0, 0);
                    var oWorksheet = Api.GetActiveSheet();
                    var ActiveCell = oWorksheet.ActiveCell;
                    ActiveCell.SetValue(Asc.scope.cell_val);
                    var characters = null;
                    var font = null;
                    Asc.scope.boldCharList.forEach(function(element, index, array) {
                        characters = ActiveCell.GetCharacters(element[0], element[1]);
                        font = characters.GetFont();
                        font.SetBold(true);
                        font.SetColor(redColor);
                    });
                    
                    var ColorCell = oWorksheet.GetCells(ActiveCell.Row, ActiveCell.Col+2);
                    ColorCell.SetValue(Asc.scope.color_cell_val);
                    Asc.scope.colorBoldCharList.forEach(function(element, index, array) {
                        characters = ColorCell.GetCharacters(element[0], element[1]);
                        font = characters.GetFont();
                        font.SetBold(true);
                    });
                    
                    ActiveCell.AutoFit(false, true);
                }, true);
            }
            }
        document.getElementById("back_button").onclick = function() {
            console.log("back button clicked");
            parts_dict = {};
            partSelectionForm.style.display = 'block';
            partConfigurationForm.style.display = 'none';
            backButton.style.display = 'none';
            status="general";
            Asc.scope.parts_dict = {};
            Asc.scope.cell_val = "";
            parts_left=0;
            part_name = "";
        }
        };

    
    window.Asc.plugin.button = function(id) {
        console.log(id);
        if (id==-1){
            this.executeCommand("close", "");
        }
    };

})(window, undefined);
