
const GOOGLE_SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIBMlIbnvEfecFbrdQ_1Wxz-2XcGn1XDswoZPwsR3_2_ChYQmcE8ubFDtft5dkJ4dqZde9xOPU5DVI/pub?output=csv";

function updateTime() {
  const now = new Date();
  document.getElementById("currentTime").innerText = now.toLocaleTimeString([], {
    hour: '2-digit',
    minute: '2-digit',
  });
}
setInterval(updateTime, 1000);
updateTime();

function fetchAndUpdateRoster() {
  Papa.parse(GOOGLE_SHEET_CSV_URL, {
    download: true,
    header: true,
    complete: function(results) {
      const rows = results.data.filter(r =>
        r["Inv 1"] === "FGZ" || r["Inv 2"] === "FGZ" || r["Res."] === "FGZ"
      );

      const now = new Date();
      let ongoing = null, next = null, allToday = [];

      // Helper functions to avoid redeclaration
      function getStartTime(row) {
        const dateParts = row["Date"].split("-");
        const [startStr] = row["Time"].split("-");
        const day = parseInt(dateParts[0], 10);
        const month = parseInt(dateParts[1], 10) - 1;
        const year = parseInt(dateParts[2], 10);
        const classDate = new Date(year, month, day);
        return new Date(`${classDate.toDateString()} ${startStr}`);
      }

      function getEndTime(row) {
        const dateParts = row["Date"].split("-");
        const [, endStr] = row["Time"].split("-");
        const day = parseInt(dateParts[0], 10);
        const month = parseInt(dateParts[1], 10) - 1;
        const year = parseInt(dateParts[2], 10);
        const classDate = new Date(year, month, day);
        return new Date(`${classDate.toDateString()} ${endStr}`);
      }

      function getClassDate(row) {
        const dateParts = row["Date"].split("-");
        const day = parseInt(dateParts[0], 10);
        const month = parseInt(dateParts[1], 10) - 1;
        const year = parseInt(dateParts[2], 10);
        return new Date(year, month, day);
      }

      rows.forEach(r => {
        const startTime = getStartTime(r);
        const endTime = getEndTime(r);
        const classDate = getClassDate(r);

        if (now >= startTime && now <= endTime) {
          ongoing = r;
        }
        if (
          now < startTime &&
          (!next || startTime < getStartTime(next))
        ) {
          next = r;
        }
        if (now.toDateString() === classDate.toDateString()) {
          allToday.push(r);
        }
      });

      document.getElementById("ongoingClass").innerHTML = ongoing ?
        `<span class='bold'>Course:</span> ${ongoing["Course/Sec"]}<br/><span class='bold'>Room:</span> ${ongoing["Room"]}` :
        "No ongoing Duty 💤";

      document.getElementById("upcomingClass").innerHTML = next ?
        `<span class='bold'>📅 ${next["Date"]}</span><br/><span class='bold'>🕒 ${next["Time"]}</span><br/>Course: ${next["Course/Sec"]}<br/><span class='bold'>Room: ${next["Room"]}</span>` :
        "No upcoming duties 🎉";

      if (allToday.length > 0) {
        const todayList = allToday.map(r =>
          `<li><span class='bold'>${r["Time"]}</span> — ${r["Course/Sec"]} (${r["Room"]})</li>`
        ).join('');
        document.getElementById("todayList").innerHTML = `<ul>${todayList}</ul>`;
      } else {
        document.getElementById("todayList").innerHTML = "<p>No duties today 🎉</p>";
      }
    }
  });
}

fetchAndUpdateRoster();
setInterval(fetchAndUpdateRoster, 300000);

document.getElementById("toggleTheme").onclick = () => {
  const isDark = document.body.classList.toggle("dark");
  document.getElementById("toggleTheme").innerText = isDark ? "☀️" : "🌙";
};
