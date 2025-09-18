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
        var phone = document.getElementById("phone");
        var city = document.getElementById("city");
        var street = document.getElementById("street");
        var house = document.getElementById("house");
        var apartment = document.getElementById("apartment");
        var floor = document.getElementById("floor");
        var entrance = document.getElementById("entrance");
        var price = document.getElementById("price");
        var dlDraft = document.getElementById("dlDraft");

        document.getElementById("buttonAddAddress").onclick = function() {
            var lift = document.querySelector('input[name="lift_group"]:checked');
            var handcaring = document.querySelector('input[name="hand_group"]:checked');
            if (handcaring.value=="true") {
                var hand_meters = document.getElementById("hand_meters");
            }
            var delivery = document.getElementById("delivery");
            Asc.scope.boldCharList = [[0,9]];
            Asc.scope.textComment = "Телефон: " + phone.value +"\n";
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,7]);
            Asc.scope.textComment = Asc.scope.textComment + "Город: " + city.value +"\n";
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,7]);
            Asc.scope.textComment = Asc.scope.textComment + "Улица: " + street.value +"\n";
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,5]);
            Asc.scope.textComment = Asc.scope.textComment + "Дом: " + house.value +"\n";
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,10]);
            Asc.scope.textComment = Asc.scope.textComment + "Квартира: " + apartment.value +"\n";
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,6]);
            Asc.scope.textComment = Asc.scope.textComment + "Этаж: " + floor.value +"\n";
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,9]);
            Asc.scope.textComment = Asc.scope.textComment + "Подъезд: " + entrance.value +"\n";
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,15]);
            Asc.scope.textComment = Asc.scope.textComment + "Грузовой лифт: " + lift.value +"\n";
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,15]);
            if(handcaring=="false"){
                    Asc.scope.textComment = Asc.scope.textComment + "Ручной пронос: Нет" +"\n";
                }
                else{
                    Asc.scope.textComment = Asc.scope.textComment + "Ручной пронос: Да, "+ hand_meters.value +" м\n";
                }
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,31]);
            Asc.scope.textComment = Asc.scope.textComment + "Озвученная стоимость доставки: " + price.value +"\n";
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,31]);
            Asc.scope.textComment = Asc.scope.textComment + "Номер черновика Деловых линий: " + dlDraft.value +"\n";
            Asc.scope.boldCharList.push([Asc.scope.textComment.length,39]);
            if(delivery=="true"){
                Asc.scope.textComment = Asc.scope.textComment + "Доставка оплачена при оформлении заказа";
            }
            window.Asc.plugin.callCommand(function() {
                var oWorksheet = Api.GetActiveSheet();
                var ActiveCell = oWorksheet.ActiveCell;
                ActiveCell.SetValue(Asc.scope.textComment);
                var characters = null;
                var font = null;
                
                Asc.scope.boldCharList.forEach(function(element, index, array) {
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
