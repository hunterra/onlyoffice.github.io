var parts_dict = {};
var status="general";
var parts_left = 0;

(function(window, undefined) {

    window.Asc.plugin.init = function() {
        document.getElementById("myForm").onsubmit = function(formObject) {
            if(status=="general"){
                var inp_elements = document.getElementsByTagName("input");
                for (let inp of inp_elements){
                    if (inp_elements[inp].checked){
                        parts_dict[inp_elements[inp].name]=[];
                    }
                }
            }
            Asc.scope.parts=''
            for (let part of parts_dict){
                Asc.scope.parts =  Asc.scope.parts + part + " ";
            }

            window.Asc.plugin.callCommand(function() {
                var oWorksheet = Api.GetActiveSheet();
                var ActiveCell = oWorksheet.ActiveCell;
                ActiveCell.SetValue(Asc.scope.parts);
            }, true);
        };
    };
    
    window.Asc.plugin.button = function(id) {
        console.log(id);
        if (id==-1){
            this.executeCommand("close", "");
        }
    };

})(window, undefined);
