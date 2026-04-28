// -------------------- CONFIG --------------------
const API_KEY = "736ce95921a5aec9007ea48fa53cfb93";
const BASE_URL = "https://api.openweathermap.org/data/2.5/weather";

// -------------------- ELEMENTS --------------------
const form = document.getElementById("weather-form");
const cityInput = document.getElementById("city-input");
const resultCard = document.getElementById("weather-result");
const errorMessage = document.getElementById("error-message");

const cityNameEl = document.getElementById("city-name");
const countryEl = document.getElementById("country");
const tempEl = document.getElementById("temp");
const iconEl = document.getElementById("icon");
const descriptionEl = document.getElementById("description");
const detailsEl = document.getElementById("details");

const timeEl = document.getElementById("current-time");

// -------------------- CITY NAME NORMALIZATION --------------------
const cityNameMap = {
    "松戸": "Matsudo",
    "まつど": "Matsudo",
    "東京": "Tokyo",
    "とうきょう": "Tokyo",
    "大阪": "Osaka",
    "おおさか": "Osaka",
    "札幌": "Sapporo",
    "さっぽろ": "Sapporo",
    "京都": "Kyoto",
    "きょうと": "Kyoto",
    "横浜": "Yokohama",
    "よこはま": "Yokohama",
};

function normalizeCityName(input) {
    const trimmed = input.trim();
    return cityNameMap[trimmed] || trimmed;
}

// -------------------- CLOCK --------------------
function updateCurrentTime() {
    const now = new Date();
    timeEl.textContent = now.toLocaleString("ja-JP");
}
setInterval(updateCurrentTime, 1000);
updateCurrentTime();

// -------------------- FETCH CURRENT WEATHER --------------------
async function getWeather(city) {
    const url = `${BASE_URL}?q=${city}&appid=${API_KEY}&units=metric&lang=ja`;

    const response = await fetch(url);
    if (!response.ok) throw new Error("都市が見つかりませんでした。");

    return response.json();
}

// -------------------- RENDER CURRENT WEATHER --------------------
function renderWeather(data) {
    console.log("Weather data:", data);
    console.log("Icon code:", data.weather[0].icon);
    cityNameEl.textContent = data.name;
    countryEl.textContent = data.sys.country;
    tempEl.textContent = `${Math.round(data.main.temp)}°C`;
    descriptionEl.textContent = data.weather[0].description;
    iconEl.src = `https://openweathermap.org/img/wn/${data.weather[0].icon}.png`;
    detailsEl.textContent = `
        体感温度: ${Math.round(data.main.feels_like)}°C  
        湿度: ${data.main.humidity}%  
        風速: ${data.wind.speed} m/s  
        気圧: ${data.main.pressure} hPa
    `;

    resultCard.classList.remove("hidden");
}

// -------------------- FORM SUBMIT HANDLER --------------------
form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const city = normalizeCityName(cityInput.value);
    try {
        errorMessage.classList.add("hidden");
        const data = await getWeather(city);
        renderWeather(data);
    } catch (error) {
        errorMessage.textContent = error.message;
        errorMessage.classList.remove("hidden");
        resultCard.classList.add("hidden");
    }
});
