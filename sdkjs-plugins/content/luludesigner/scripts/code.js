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
            var bold_char_list = [[0,4]];
            Asc.scope.textComment = Asc.scope.textComment + "Телефон: " + phone.value +"\n";
            bold_char_list.push([Asc.scope.textComment.length,9]);
            Asc.scope.textComment = Asc.scope.textComment + "e-mail: " + email.value +"\n";
            bold_char_list.push([Asc.scope.textComment.length,8]);
            Asc.scope.textComment = Asc.scope.textComment + "Ссылка на сайт/inst: " + site.value +"\n";
            bold_char_list.push([Asc.scope.textComment.length,21]);
            Asc.scope.textComment = Asc.scope.textComment + "Условия: " + conditions.value +"\n";
            bold_char_list.push([Asc.scope.textComment.length,9]);
            Asc.scope.textComment = Asc.scope.textComment + "Сумма Б: " + sum.value +"\n";
            bold_char_list.push([Asc.scope.textComment.length,9]);
            
            console.log(Asc.scope.textComment);
            window.Asc.plugin.callCommand(function() {
                var oWorksheet = Api.GetActiveSheet();
                var ActiveCell = oWorksheet.ActiveCell;
                ActiveCell.SetValue(Asc.scope.textComment);
                var characters = null;
                var font = null;
                
                bold_char_list.forEach(function(element, index, array) {
                    characters = ActiveCell.GetCharacters(element[0], element[1]);
                    font = characters.GetFont();
                    font.SetBold(true);
                });

            }, true);
        };
    };
    
    window.Asc.plugin.button = function() {
        this.executeCommand("close", "");
    };

})(window, undefined);
