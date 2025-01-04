// 작성자 리스트
const authors = ["정명자", "유숙재", "박지예", "문일지", "김성혁", "김도경", "윤진", "노해진", "표세흠"];

// 현재 년도와 월 추적
let currentYear = 2025;
let currentMonth = 0;
let selectedDate = ""; // 선택된 날짜

// 인수인계 항목 리스트
const handoverItemsList = [
    "공통사항",
    "수란씨네",
    "보용씨네",
    "주공 201동",
    "주공 301동",
    "주공 303동",
    "주공 308동",
    "황전 103동",
    "황전 104동",
    "덕산 103동",
    "덕산 104동",
    "전일 야간 인계 - 1",
    "전일 야간 인계 - 2"
];

// 데이터 저장소 (각 날짜별 데이터 저장)
const handoverData = {};

// 요일 표시
const weekdays = ["일", "월", "화", "수", "목", "금", "토"];

// 달력 생성 함수
function generateCalendar() {
    const calendarContainer = document.getElementById("calendar-container");
    const monthYear = document.getElementById("month-year");

    // 현재 월 표시
    monthYear.innerText = `${currentYear}년 ${currentMonth + 1}월`;

    // 달력 초기화
    calendarContainer.innerHTML = "";

    // 요일 표시
    weekdays.forEach((day, index) => {
        const weekdayDiv = document.createElement("div");
        weekdayDiv.className = "weekday";
        weekdayDiv.innerText = day;

        if (index === 0) weekdayDiv.style.color = "red"; // 일요일
        if (index === 6) weekdayDiv.style.color = "blue"; // 토요일

        calendarContainer.appendChild(weekdayDiv);
    });

    // 해당 월 첫날과 마지막 날짜 계산
    const firstDay = new Date(currentYear, currentMonth, 1).getDay();
    const lastDate = new Date(currentYear, currentMonth + 1, 0).getDate();

    // 빈 칸 추가
    for (let i = 0; i < firstDay; i++) {
        const emptyDiv = document.createElement("div");
        emptyDiv.className = "calendar-day empty";
        calendarContainer.appendChild(emptyDiv);
    }

    // 날짜 추가
    for (let i = 1; i <= lastDate; i++) {
        const dayDiv = document.createElement("div");
        dayDiv.className = "calendar-day";
        dayDiv.innerText = i;

        // 주말 스타일 추가
        const dayOfWeek = new Date(currentYear, currentMonth, i).getDay();
        if (dayOfWeek === 0) dayDiv.style.color = "red"; // 일요일
        if (dayOfWeek === 6) dayDiv.style.color = "blue"; // 토요일

        dayDiv.addEventListener("click", () => {
            updateSelectedDate(i);
            displayHandoverItems();
        });

        calendarContainer.appendChild(dayDiv);
    }
}

// 날짜 선택 업데이트
function updateSelectedDate(day) {
    selectedDate = `${currentYear}-${String(currentMonth + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    const handoverTitle = document.getElementById("handover-title");
    handoverTitle.innerText = `${currentYear}년 ${currentMonth + 1}월 ${day}일 / 인수인계`;
}

// 인수인계 항목 표시
function displayHandoverItems() {
    const handoverContent = document.getElementById("handover-content");
    handoverContent.innerHTML = ""; // 기존 항목 삭제

    // 현재 날짜의 데이터 가져오기 (없으면 빈 데이터 생성)
    if (!handoverData[selectedDate]) {
        handoverData[selectedDate] = handoverItemsList.map(item => ({
            항목: item,
            인계내용: "",
            작성자: ""
        }));
    }

    // 데이터 표시
    handoverData[selectedDate].forEach((itemData, index) => {
        const row = document.createElement("tr");

        const cell1 = document.createElement("td");
        cell1.innerText = itemData.항목;

        const cell2 = document.createElement("td");
        const textarea = document.createElement("textarea");
        textarea.className = "dynamic-textarea";
        textarea.placeholder = "인계 내용을 적어주세요.";
        textarea.value = itemData.인계내용; // 저장된 내용 로드
        textarea.addEventListener("input", () => {
            handoverData[selectedDate][index].인계내용 = textarea.value; // 데이터 저장
        });
        cell2.appendChild(textarea);

        const cell3 = document.createElement("td");
        const select = document.createElement("select");
        const defaultOption = document.createElement("option");
        defaultOption.text = "작성자 선택";
        select.appendChild(defaultOption);

        authors.forEach(author => {
            const option = document.createElement("option");
            option.value = author;
            option.innerText = author;
            if (itemData.작성자 === author) {
                option.selected = true; // 저장된 작성자 로드
            }
            select.appendChild(option);
        });

        select.addEventListener("change", () => {
            handoverData[selectedDate][index].작성자 = select.value; // 데이터 저장
        });
        cell3.appendChild(select);

        row.appendChild(cell1);
        row.appendChild(cell2);
        row.appendChild(cell3);

        handoverContent.appendChild(row);

        // 텍스트 높이 자동 조정
        adjustTextareaHeight(textarea);
    });
}

// 자동 크기 조정 함수
document.addEventListener("input", (event) => {
    if (event.target.classList.contains("dynamic-textarea")) {
        adjustTextareaHeight(event.target);
    }
});

function adjustTextareaHeight(textarea) {
    textarea.style.height = "auto";
    textarea.style.height = `${textarea.scrollHeight}px`;
}

// 엑셀로 저장하기 기능
function saveToExcel() {
    const data = handoverData[selectedDate] || []; // 현재 날짜 데이터 가져오기

    // 선택된 날짜를 데이터 맨 위에 추가
    const formattedDate = selectedDate || "선택된 날짜 없음";
    const excelData = [{ 항목: "날짜", 인계내용: formattedDate, 작성자: "" }, ...data];

    // 줄바꿈 옵션이 적용된 데이터 생성
    const worksheet = XLSX.utils.json_to_sheet(excelData);

    // 줄바꿈 스타일 설정
    Object.keys(worksheet).forEach(cell => {
        if (cell[0] === '!') return; // 메타데이터 제외

        // 각 셀에 줄바꿈 옵션 추가
        worksheet[cell].s = {
            alignment: {
                wrapText: true, // 줄바꿈 활성화
                vertical: "top", // 텍스트를 위쪽 정렬
                horizontal: "left" // 텍스트를 왼쪽 정렬
            }
        };
    });

    // 워크북 생성
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "인수인계");

    // 엑셀 파일 다운로드
    XLSX.writeFile(workbook, `${selectedDate} 인수인계.xlsx`);
}

// 이전/다음 월 버튼 클릭 이벤트
document.getElementById("prev-month-btn").addEventListener("click", () => {
    currentMonth--;
    if (currentMonth < 0) {
        currentMonth = 11;
        currentYear--;
    }
    generateCalendar();
});

document.getElementById("next-month-btn").addEventListener("click", () => {
    currentMonth++;
    if (currentMonth > 11) {
        currentMonth = 0;
        currentYear++;
    }
    generateCalendar();
});

// 초기 달력 생성
generateCalendar();

// 엑셀 다운로드 버튼 이벤트 추가
document.getElementById("download-btn").addEventListener("click", saveToExcel);