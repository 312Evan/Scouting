const express = require('express');
const path = require('path');
const fs = require('fs');
const axios = require('axios');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');

require('dotenv').config();

const app = express();
const PORT = 3000;
const TBA_KEY = process.env.BLUE_ALLIANCE;

app.use(bodyParser.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));
app.disable('view cache');

const usersPath = path.join(__dirname, 'data', 'users.json');

app.get('/', async (req, res) => {
    try {
        const currentYear = new Date().getFullYear();

        const response = await axios.get(
            `https://www.thebluealliance.com/api/v3/team/frc2068/events/${currentYear}`,
            { headers: { 'X-TBA-Auth-Key': TBA_KEY } }
        );

        res.render('index', { events: response.data });
    } catch (err) {
        console.error(err.response?.data || err.message);
        res.send("Error loading events.");
    }
});


app.post('/select-event', async (req, res) => {
    const { eventKey } = req.body;
    const schedulePath = path.join(__dirname, 'data', `schedule_${eventKey}.json`);
    const usersData = JSON.parse(fs.readFileSync(usersPath));

    try {
        if (fs.existsSync(schedulePath)) {

            let eventName = eventKey;
            try {
                const eventRes = await axios.get(
                    `https://www.thebluealliance.com/api/v3/event/${eventKey}`,
                    { headers: { 'X-TBA-Auth-Key': TBA_KEY } }
                );
                eventName = eventRes.data.name || eventKey;
            } catch {}

            return res.render('scouterSelect', {
                eventKey,
                eventName,
                users: usersData.users,
                scheduleAvailable: true
            });
        }

        const scouters = Object.keys(usersData.users)
            .filter(u => usersData.users[u].scouter);

        if (scouters.length < 1) {
            return res.send("Need at least 1 scouter to assign matches.");
        }

        const matchesRes = await axios.get(
            `https://www.thebluealliance.com/api/v3/event/${eventKey}/matches`,
            { headers: { 'X-TBA-Auth-Key': TBA_KEY } }
        );

        const matches = matchesRes.data
            .filter(m => m.comp_level === "qm")
            .sort((a, b) => a.match_number - b.match_number);

        if (!matches.length) {
            return res.send("No qualification matches found.");
        }

        let schedule = {};
        let assignmentCounts = {};
        const MAX_MATCHES_PER_SCOUTER = 16;

        scouters.forEach(s => {
            schedule[s] = [];
            assignmentCounts[s] = 0;
        });

        matches.forEach(match => {
            const teams = [
                ...(match.alliances?.red?.team_keys || []),
                ...(match.alliances?.blue?.team_keys || [])
            ];

            teams.forEach(team => {

                const eligible = scouters.filter(
                    s => assignmentCounts[s] < MAX_MATCHES_PER_SCOUTER
                );

                if (!eligible.length) return;

                const min = Math.min(...eligible.map(s => assignmentCounts[s]));
                const candidates = eligible.filter(s => assignmentCounts[s] === min);

                const chosen = candidates[Math.floor(Math.random() * candidates.length)];

                schedule[chosen].push({
                    match: match.match_number,
                    team
                });

                assignmentCounts[chosen]++;
            });
        });

        fs.writeFileSync(schedulePath, JSON.stringify(schedule, null, 4));

        let eventName = eventKey;
        try {
            const eventRes = await axios.get(
                `https://www.thebluealliance.com/api/v3/event/${eventKey}`,
                { headers: { 'X-TBA-Auth-Key': TBA_KEY } }
            );
            eventName = eventRes.data.name || eventKey;
        } catch {}

        res.render('scouterSelect', {
            eventKey,
            eventName,
            users: usersData.users,
            scheduleAvailable: true
        });

    } catch (err) {
        console.error(err.response?.data || err.message);
        res.send("Error generating schedule.");
    }
});

app.post('/login', async (req, res) => {
    const { eventKey, username } = req.body;
    const schedulePath = path.join(__dirname, 'data', `schedule_${eventKey}.json`);

    if (!fs.existsSync(schedulePath)) {
        return res.send("Schedule not generated yet.");
    }

    const schedule = JSON.parse(fs.readFileSync(schedulePath));
    let assignedMatches = schedule[username] || [];

    assignedMatches.sort((a, b) => {
        if (a.match === b.match) {
            return Number(a.team.replace('frc', '')) -
                   Number(b.team.replace('frc', ''));
        }
        return Number(a.match) - Number(b.match);
    });

    let eventName = eventKey;
    try {
        const eventRes = await axios.get(
            `https://www.thebluealliance.com/api/v3/event/${eventKey}`,
            { headers: { 'X-TBA-Auth-Key': TBA_KEY } }
        );
        eventName = eventRes.data.name || eventKey;
    } catch {}

    res.render('dashboard', {
        eventKey,
        eventName,
        username,
        matches: assignedMatches
    });
});

app.post('/scout', (req, res) => {
    const { eventKey, username, matchNumber, teamNumber } = req.body;
    res.render('scoutForm', { eventKey, username, matchNumber, teamNumber });
});


app.post('/submit-scout', async (req, res) => {
    const data = req.body;
    const { eventKey, username, matchNumber, teamNumber } = data;

    const exportPath = path.join(__dirname, 'exports', `${eventKey}.xlsx`);
    const workbook = new ExcelJS.Workbook();

    if (fs.existsSync(exportPath)) {
        await workbook.xlsx.readFile(exportPath);
    }

    let sheet = workbook.getWorksheet("Summary");
    if (!sheet) {
        sheet = workbook.addWorksheet("Summary");
        sheet.addRow(Object.keys(data));
    }

    sheet.addRow(Object.values(data));
    await workbook.xlsx.writeFile(exportPath);

    const schedulePath = path.join(__dirname, 'data', `schedule_${eventKey}.json`);
    const schedule = JSON.parse(fs.readFileSync(schedulePath));

    if (schedule[username]) {
        schedule[username] = schedule[username].filter(item =>
            !(String(item.match) === String(matchNumber) &&
              String(item.team) === String(teamNumber))
        );
    }

    fs.writeFileSync(schedulePath, JSON.stringify(schedule, null, 4));

    const users = JSON.parse(fs.readFileSync(usersPath));
    users.users[username].shiftsSubmitted =
        (users.users[username].shiftsSubmitted || 0) + 1;

    fs.writeFileSync(usersPath, JSON.stringify(users, null, 4));

    let eventName = eventKey;
    try {
        const eventRes = await axios.get(
            `https://www.thebluealliance.com/api/v3/event/${eventKey}`,
            { headers: { 'X-TBA-Auth-Key': TBA_KEY } }
        );
        eventName = eventRes.data.name || eventKey;
    } catch {}

    res.render('submitSuccess', { username, eventName });
});


app.listen(PORT, () =>
    console.log(`Server running on http://localhost:${PORT}`)
);