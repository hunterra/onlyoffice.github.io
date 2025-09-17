/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
(function(window, undefined) {

    window.Asc.plugin.init = function() {
        var fio = document.getElementById("fio");
        var phone = document.getElementById("phone");
        var email = document.getElementById("email");
        var site = document.getElementById("site");
        var sum = document.getElementById("sum");

        document.getElementById("buttonAddDesign").onclick = function() {
            var conditions = document.querySelector('input[name="cond_group"]:checked');
            if (conditions.value=="Иное") {
                conditions = document.getElementById("custom_cond");
            }
            Asc.scope.textComment = "ФИО: " + fio.value +"\n";
            Asc.scope.textComment = Asc.scope.textComment + "Телефон: " + phone.value +"\n";
            Asc.scope.textComment = Asc.scope.textComment + "e-mail: " + email.value +"\n";
            Asc.scope.textComment = Asc.scope.textComment + "Ссылка на сайт/inst: " + site.value +"\n";
            Asc.scope.textComment = Asc.scope.textComment + "Условия: " + conditions.value +"\n";
            Asc.scope.textComment = Asc.scope.textComment + "Сумма Б: " + sum.value +"\n";
            
            console.log(Asc.scope.textComment);
            window.Asc.plugin.callCommand(function() {
                var oWorksheet = Api.GetActiveSheet();
                var ActiveCell = oWorksheet.ActiveCell;
                let cellContent = ActiveCell.GetContent(); // Get the content of the cell
                // Now you can work with the cellContent, for example, to add text:
                let paragraph = Api.CreateParagraph();
                let run = Api.CreateRun();
                run.AddText("This is just a sample text. ");
                paragraph.AddElement(run);
                run = Api.CreateRun();
                run.SetFontFamily("Comic Sans MS");
                run.AddText("This is a text run with the font family set to 'Comic Sans MS'.");
                paragraph.AddElement(run);
                cellContent.Push(paragraph);
            }, true);
        };
    };
    
    window.Asc.plugin.button = function() {
        this.executeCommand("close", "");
    };

})(window, undefined);
