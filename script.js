let globalData;

document.getElementById('excel-file').addEventListener('change', handleFileSelect, false);
document.getElementById('generate-chart').addEventListener('click', () => createChart(globalData), false);

function handleFileSelect(event) {
    // 獲取選擇的文件
    const file = event.target.files[0];
    // 創建讀取文件的物件
    const reader = new FileReader();
    // 設置讀取完成後的回調函數
    reader.onload = function (e) {
        // 獲取文件內容 data = e.target.result 表示獲取讀取到的文件內容
        const data = e.target.result;
        // workbook 就是整份 Excel 文件 data 就是文件內容 這裡使用 XLSX 來解析文件 並將文件解析為 JSON 格式 type: 'binary' 表示文件的類型是二進制
        const workbook = XLSX.read(data, { type: 'binary' });
        // 獲取 Excel 文件中第一張表格的名稱
        const sheetName = workbook.SheetNames[0];
        // 獲取第一張表格內容
        const worksheet = workbook.Sheets[sheetName];
        // 將表格內容轉換為 JSON 格式
        globalData = XLSX.utils.sheet_to_json(worksheet);
    };
    //readAsBinaryString 將文件讀取為二進制字串
    reader.readAsBinaryString(file);
}

// 創建柱狀圖
function createChart(data) {
    // 如果沒有選擇文件則提示用戶
    if (!data) {
        alert('請先選擇一個 Excel 文件');
        return;
    }

    // 將數據處理為需要的格式 data.reduce() 將數據轉換為對象 { '產品名稱': '委工重量' }
    // reduce() 方法將一個累加器及陣列中每項元素（由左至右）傳入回呼函式，將陣列化為單一值。
    let chartData = data.reduce((acc, cur) => {
        // 讓產品名稱作為鍵 委工重量作為值
        let product = cur['產品名稱'];
        let weight = cur['委工重量'];
        // parseFloat() 函式可解析一個字串，並回傳一個浮點數。
        weight = parseFloat(weight);
        // 如果產品名稱不存在則初始化為0 如果存在則將委工重量相加
        acc[product] = (acc[product] || 0) + weight;
        // 返回累加器
        return acc;
        // {} 為初始值
    }, {});

    // 獲取前10個產品 Object.entries() 方法會回傳一個給定物件自身可列舉屬性的鍵值對陣列，其排列順序和使用 for...in 迴圈遍歷該物件時一致（兩者的差異在於 for-in 迴圈還會列舉其原型鍊中的屬性）。
    chartData = Object.entries(chartData)
        // sort() 方法會原地（in place）對一個陣列的所有元素進行排序，並回傳此陣列。排序不一定是穩定的。預設排序順序是根據字串 Unicode 編碼位置而定。 
        //.sort((a, b) => b[1] - a[1]) 代表陣列中的第二個元素由大到小排序
        // .slice(0, 10) 代表取前10個元素 用來顯示前10個產品 可以獲得前10個產品的鍵值對陣列
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    // 分割鍵和值  map() 方法會建立一個新陣列，其結果是該陣列中的每個元素都會呼叫一次提供的函式後所回傳的結果。
    //  chartData.map(item => item[0]) 代表取出鍵 chartData.map(item => item[1]) 代表取出值
    const labels = chartData.map(item => item[0]);
    const values = chartData.map(item => item[1]);

    // 創建柱狀圖 new Chart是Chart.js的方法
    new Chart(document.getElementById('bar-chart'), {
        // type: 'bar' 代表創建柱狀圖
        type: 'bar',
        // data: {} 代表圖表的數據
        data: {
            // labels: labels 代表圖表的標籤
            labels: labels,
            // datasets: [] 代表圖表的數據集
            datasets: [{
                // label: '委工重量' 代表數據集的標籤
                label: '委工重量',
                // data: values 代表數據集的數據
                data: values,
                backgroundColor: 'rgba(0, 123, 255, 0.5)'
            }]
        },
        // options: {} 代表圖表的配置
        options: {
            //scales: {} 代表圖表的刻度 yAxes: [] 代表 y 軸的刻度 ticks: {} 代表刻度的配置 beginAtZero: true 代表從0開始 這樣就不會出現負數
            scales: {
                // yAxes: [] 代表 y 軸的刻度
                yAxes: [{ ticks: { beginAtZero: true } }]
            }
        }
    });
}
