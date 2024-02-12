let DATA = undefined;
let DATA_VIEW = undefined;
let textSearch = {};
let selectedRange = '';
let selectInputRange = document.querySelector('#importRange');
let selectInputSearch = document.querySelector('#importSelectSearchCard');
let selectInputSearchCode = document.querySelector('#importSelectSearchCardCode');
let inputSeacrhCard = document.querySelector('#importSeacrhCard');
let inputSeacrhCardCode = document.querySelector('#importSeacrhCardCode');
let inputBeforeText = document.querySelector('#importBeforeText');
let inputAfterText = document.querySelector('#importAfterText');
let form = document.getElementById('form-data-import');
let ViewDataTable = document.querySelector('#view-data-table tbody');
let viewCountAll = document.querySelector('#viewCountAll');
let viewCountDone = document.querySelector('#viewCountDone');
let viewCountNotDone = document.querySelector('#viewCountNotDone');
let StateScan = {
    timer: null,
    text: '',
    isFocus: true
}
const ModalSettings = document.getElementById('modal_settings');
let context_menu = {
    state:{
        isVisible: false,
        selected: null
    },
    el: document.querySelector('#context_menu')
};



document.querySelector('#importFiles').addEventListener('change',async(e)=>{
    if(e.target.files[0] == undefined) return;
    DATA = undefined;
    textSearch = {};
    selectedRange = '';
    let xlsx = await readXLSX(e);

    document.querySelector('#table_import tbody').innerHTML = xlsx.tbodyHTML;
    document.querySelectorAll('#table_import tbody tr').forEach((e,i)=>{if(i>100)e.style.display='none'});
    document.querySelector('#table_import .table-sticky').innerHTML = xlsx.column_name.map(e=>(`<th scope="col">${e}</th>`)).join('');
    selectInputRange.innerHTML = '<option value="" selected>Не выбран</option>'+xlsx.column_name.map(e=>(`<option value="${e}">${e}:${e}</option>`)).join('')
    DATA = getData(document.querySelector('#table_import tbody'));
    xlsx = null;
})

async function readXLSX(files){
    const file = files.target.files[0];
    const data = await file.arrayBuffer();

    const wb = XLSX.read(data);
    const ws = wb.Sheets[wb.SheetNames[0]];
    
    let htmlTable = document.createElement('html');
    htmlTable.innerHTML = XLSX.utils.sheet_to_html(ws, { id: "TableSheets" });
    let tbody = htmlTable.querySelector('tbody')
    return {
        column_name: Object.keys(getData(tbody)).sort(),
        tbodyHTML: tbody.innerHTML
    }
}

function getData(tbody) {
    let DATA = {}
    tbody.querySelectorAll('td').forEach(e=>{
        let word = e.getAttribute('id').replace('TableSheets-','').match(/([A-Z]+)/g);
        if(DATA[word] == undefined){
            DATA[word] = [e];
            return;
        }
        DATA[word].push(e)
    })
    return DATA;
}

document.querySelector('#table_import tbody').addEventListener('mousedown',(e)=>{
    if(DATA == undefined) return;
    let range = e.target.getAttribute('id').replace('TableSheets-','').match(/([A-Z]+)/g);
    selectColumns(range);
})

document.querySelector('#table_import thead tr').addEventListener('mousedown',(e)=>{
    if(DATA == undefined) return;
    let range = e.target.textContent;
    selectColumns(range);
})


function selectColumns(range){
    if(selectedRange == range) return;
    selectedRange = range;
    selectInputRange.value = range;
    document.querySelectorAll('.table-active').forEach(e=>e.className = '');
    if(selectedRange == '') return;
    DATA[range].map(e=>e.className = 'table-active');
    if(textSearch[range] != undefined) return;
    textSearch[range] = DATA[range].map(e=>e.textContent).join('\n')
}


function convertStringToRegex(regexString) {
    let match = regexString.match(new RegExp('^/(.*?)/([gimy]*)$'));
    let regex = new RegExp(match[1], match[2]);
    return regex;
}

selectInputRange.addEventListener('change',e=>{
    selectColumns(e.target.value)
})

inputSeacrhCard.addEventListener('change',e=>{
    if(e.target.value != selectInputSearch.value || e.target.value == '') {
        selectInputSearch.value = '';
        return;
    }
})

inputSeacrhCardCode.addEventListener('change',e=>{
    if(e.target.value != selectInputSearchCode.value || e.target.value == '') {
        selectInputSearchCode.value = '';
        return;
    }
})

selectInputSearchCode.addEventListener('change',e=>inputSeacrhCardCode.value = e.target.value)
selectInputSearch.addEventListener('change',e=>inputSeacrhCard.value = e.target.value)

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}


form.addEventListener('submit', async(e)=>{
    e.preventDefault();
    let data = new FormData(e.target);
    let range = data.get('range');
    let seacrhCard = data.get('seacrhCard');
    let seacrhCardCode = data.get('seacrhCardCode');
    let beforeText = data.get('beforeText');
    let afterText = data.get('afterText');
    let cards = null;
    let codeCards = null;
    let cardsText = '';
    try{
        cards = textSearch[range].match(convertStringToRegex(seacrhCard));
        if(cards == null){
            alert('Ничего не найдено при поиске карт!');
            return;
        }
        codeCards = cards.join('\n').match(convertStringToRegex(seacrhCardCode));
        if(codeCards == null){
            alert('Ничего не найдено при поиске кода карты!');
            return;
        }
        cardsText = cards.map(e=>(beforeText+e+afterText));
        
    }catch{
        if(DATA == undefined) {alert('Выберите файл!!'); return;}
        alert('Ошибка регулярных выражений!')
        return;
    }
    document.querySelector('#submit_import').disabled = true;
    document.querySelector('#submit_import').innerHTML = `<span class="spinner-border spinner-border-sm" aria-hidden="true"></span>
    <span role="status">Loading...</span>`;
    let ask = false;
    await delay(100).then(()=>{
        ask = confirm(`Вот что получилось:\nНомер карты: ${cards[0]}\nКод карты: ${codeCards[0]}\nТекст печати: ${cardsText[0]}\nПравильно?`)
    })
    if(ask){
        if (selectInputSearch.value == '/([0-9]*-[0-9]*-[0-9]*\\/[0-9]*)/g') {
            const double = {};
            cards.map((data, i) => {
                let indx = data.split('-')[0]+data.split('-')[1];
                double[indx] = double[indx]==undefined? [i]:[...double[indx], i]
            });
            let check_double = Object.entries(double).filter(e=>e[1].length>1).map(e=>cards[e[1][0]]);
            if(check_double.length !=0) {
                alert(`В вашем списке номеров обнаруженны дубликаты:\n${check_double.join('\n')}`);
                document.querySelector('#submit_import').disabled = false;
                document.querySelector('#submit_import').innerHTML = `Применить`;
                return
            }
        }

        if(selectInputSearchCode.value == '/([0-9]*\\/[0-9]*)/g'){
            codeCards = codeCards.map((e)=>{
                let code = e;
                if (code.split('/')[1].length!=5) {
                    code = e.split('/')[0]+'/'+String('0').repeat(5-e.split('/')[1].length)+e.split('/')[1];
                }
                return code;
            })
        }
        let out = cards.map((e,i)=>({
            cards:e,
            codeCards: codeCards[i],
            printCards: cardsText[i]
        }))
        ViewDataTable.innerHTML = '';
        out.map((e,i)=>{
            ViewDataTable.innerHTML += `
            <tr class="" id="tr-${i}" data-tr-index="${i}">
                <th scope="row">${i+1}</th>
                <td>${e.cards}</td>
                <td>${e.codeCards}</td>
                <td>${e.printCards}</td>
            </tr>
            `;
        })
        DATA_VIEW = out;
        viewCountAll.textContent = out.length;
        viewCountDone.textContent = '0';
        viewCountNotDone.textContent = out.length;
        document.querySelector('#alert-success').style.opacity = 1;
        setTimeout(()=>{document.querySelector('#alert-success').style.opacity = 0;},2000);
    }
    document.querySelector('#submit_import').disabled = false;
    document.querySelector('#submit_import').innerHTML = `Применить`;

})


document.addEventListener('keydown', (e)=>{
    if(!StateScan.isFocus) return;
    if (e.keyCode === 32 || e.code === '' || e.shiftKey) {  
        e.preventDefault();  
    } 
    StateScan.text += e.key;
    checkScan();
    clearTimeout(StateScan.timer);
    StateScan.timer = setTimeout(()=>{
        document.querySelector("#codeCard").textContent = StateScan.text;
        if(DATA_VIEW != undefined) findCode(StateScan.text.replaceAll(' ',''));
        StateScan.text='';
    },200)
    
})



function findCode(code) {
    let indx = DATA_VIEW.findIndex((element) => element.codeCards == code);
    if(indx == -1) {
        document.querySelector("#codeSCUD").textContent = 'Нет данных';
        document.querySelector("#numberDoc").textContent = 'Нет данных';
        document.querySelector("#statusView").textContent = 'Отсутсвует';
        document.querySelector("#statusView").style.color = 'red';
        return;
    }

    document.querySelector("#codeSCUD").textContent = DATA_VIEW[indx].codeCards;
    document.querySelector("#numberDoc").textContent = DATA_VIEW[indx].cards;
    document.querySelector("#statusView").textContent = 'Полученно';
    document.querySelector("#statusView").style.color = 'green';
    document.getElementById(`tr-${indx}`).scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
    setSelected(ViewDataTable.querySelector(`#tr-${indx}`));
}

function checkScan(){
    document.querySelector("#codeCard").textContent = 'Нет данных';
    document.querySelector("#codeSCUD").textContent = 'Нет данных';
    document.querySelector("#numberDoc").textContent = 'Нет данных';
    document.querySelector("#statusView").textContent = 'Чтение';
    document.querySelector("#statusView").style.color = '#000';
}

ModalSettings.addEventListener('hide.bs.modal', () => StateScan.isFocus = true);
ModalSettings.addEventListener('show.bs.modal', () => StateScan.isFocus = false);

document.querySelector('#printXLSX').addEventListener('click', ()=>{
    let db = DATA_VIEW.map(d=>(d.printCards))
    let print = db.map((e,i)=>({print:e, printReverse:db[(db.length-1)-i]}))
    let ws = XLSX.utils.json_to_sheet(print);
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Dates");
    XLSX.writeFile(wb, `PrinterDB_${new Date().getTime()}.xlsx`, { compression: true });
})

context_menu.el.addEventListener('click', function(e){
    e.preventDefault();
    let event = e.target.getAttribute('data-event')
    if(event == undefined || event == null) return;
    eval(`${event}(context_menu.state.selected)`);
})


function setSelected(el){
    if(el.className == 'table-success') return;
    el.className = 'table-success';
    viewCountDone.textContent = +viewCountDone.textContent+1;
    viewCountNotDone.textContent = +viewCountNotDone.textContent-1;
}
function setNotSelected(el){
    if(el.className == '') return;
    el.className = '';
    viewCountDone.textContent = +viewCountDone.textContent-1;
    viewCountNotDone.textContent = +viewCountNotDone.textContent+1;
}

document.querySelector('#view-data-table tbody').addEventListener('contextmenu',(e)=>{
    e.preventDefault();
    if(DATA_VIEW == undefined) return;
    context_menu.el.style.top = e.clientY + 'px';
    context_menu.el.style.left = e.clientX + 'px';
    context_menu.el.style.display = 'block';
    context_menu.state.isVisible = true;
    context_menu.state.selected = e.target.parentNode
})

document.addEventListener('click', ()=>{
    if(context_menu.state.isVisible){
        context_menu.el.style.display = 'none';
        context_menu.state.isVisible = false;
        context_menu.state.selected = null;
    }
})
let btnSaveData = document.querySelector('#saveData');
let localStorageIsSave = false;
btnSaveData.addEventListener('click',(e)=>{
    if(localStorageIsSave){
        localStorageIsSave = false;
        localStorage.removeItem('saveData');
        e.target.textContent = 'Сохранить в память'
        return;
    }
    e.target.textContent = 'Очистить память';
    localStorageIsSave = true;
    localStorage.setItem('saveData',JSON.stringify({
        DATA_VIEW,
        textSearch,
        selectedRange,
        table_settings: document.querySelector('#table_import tbody').innerHTML,
        table_settings_columns: document.querySelector('#table_import .table-sticky').innerHTML,
        table_view: ViewDataTable.innerHTML,
        selectInputRange: selectInputRange.innerHTML,
        inputSeacrhCard: inputSeacrhCard.value,
        inputSeacrhCardCode: inputSeacrhCardCode.value,
        inputBeforeText: inputBeforeText.value,
        inputAfterText: inputAfterText.value,
        selectInputSearch: selectInputSearch.value,
        selectInputSearchCode: selectInputSearchCode.value
    }));
})


function importSaveFile(json){
    let data = JSON.parse(json)
    DATA_VIEW = data.DATA_VIEW;
    textSearch = data.textSearch;
    selectedRange = data.selectedRange;
    document.querySelector('#table_import tbody').innerHTML = data.table_settings;
    document.querySelector('#table_import .table-sticky').innerHTML = data.table_settings_columns;
    ViewDataTable.innerHTML = data.table_view;
    selectInputRange.innerHTML = data.selectInputRange;
    selectInputRange.value = data.selectedRange;
    inputSeacrhCard.value = data.inputSeacrhCard;
    inputSeacrhCardCode.value = data.inputSeacrhCardCode;
    inputBeforeText.value = data.inputBeforeText;
    inputAfterText.value = data.inputAfterText;
    selectInputSearch.value = data.selectInputSearch;
    selectInputSearchCode.value = data.selectInputSearchCode;
    DATA = getData(document.querySelector('#table_import tbody'));
    viewCountAll.textContent = data.DATA_VIEW.length
    viewCountDone.textContent = document.querySelectorAll('.table-success').length;
    viewCountNotDone.textContent = viewCountAll.textContent - viewCountDone.textContent
}


(function(){
    let json = localStorage.getItem('saveData');
    if(json !== null){
        localStorageIsSave = true;
        btnSaveData.textContent = 'Очистить память';
        importSaveFile(json);
        json = '';
    }
})();

