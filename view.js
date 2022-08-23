let sheetID = ['15ci4qvWgF1SE_9EmHRNLrFKdZC06xZRzvo6xNi2T2lA', '1h8enY_zOwImhRob6YcNDqZJxPA5YCvSSIDXjd1NC4mc','1KajHTKsGEK8VbhMIGNRQUtwXL5tWr5HUKFcBeUSbtmg']; 
let id = [sheetID.length];
let sheetList = [1,2,3];
let sheet = "";
let OPTION = "";
const SELECT = document.getElementById('selected');
     sheetList.forEach((id, idx) => {
        OPTION = document.createElement('option');
        OPTION.value =sheetID;
        OPTION.innerText = `${idx + 1}번째 시트`;
        OPTION.id = `sheet${idx + 1}`;
        SELECT.appendChild(OPTION);
     });   