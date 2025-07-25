
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
  Papa.parse("https://docs.google.com/spreadsheets/d/e/2PACX-1vTIBMlIbnvEfecFbrdQ_1Wxz-2XcGn1XDswoZPwsR3_2_ChYQmcE8ubFDtft5dkJ4dqZde9xOPU5DVI/pub?output=csv", {
    download: true,
    header: true,
    complete: function(results) {
      const rows = results.data.filter(r =>
        r["Inv 1"] === "FGZ" || r["Inv 2"] === "FGZ" || r["Res."] === "FGZ"
      );

      const now = new Date();
      let ongoing = null, next = null, allToday = [];

      rows.forEach(r => {
        const dateParts = r["Date"].split("-");
        const [startStr, endStr] = r["Time"].split("-");
        const classDate = new Date(`${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`);
        const startTime = new Date(`${classDate.toDateString()} ${startStr}`);
        const endTime = new Date(`${classDate.toDateString()} ${endStr}`);

        if (now >= startTime && now <= endTime) {
          ongoing = r;
        }
        if (now < startTime && (!next || startTime < new Date(`${next["Date"].split("-")[2]}-${next["Date"].split("-")[1]}-${next["Date"].split("-")[0]} ${next["Time"].split("-")[0]}`))) {
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
