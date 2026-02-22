// server.js
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

const usersPath = path.join(__dirname, 'data', 'users.json');

app.get('/', async (req, res) => {
    try {
        // const currentYear = new Date().getFullYear();
        const currentYear = 2025;
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

    try {
        const usersData = JSON.parse(fs.readFileSync(usersPath));
        const scouters = Object.keys(usersData.users).filter(u => usersData.users[u].scouter);

        if (scouters.length < 1) {
            return res.send("Need at least 1 scouter to assign matches.");
        }

        const matchesRes = await axios.get(
            `https://www.thebluealliance.com/api/v3/event/${eventKey}/matches`,
            { headers: { 'X-TBA-Auth-Key': TBA_KEY } }
        );

        const matches = matchesRes.data.filter(m => m.comp_level === "qm");

        if (!matches || matches.length === 0) {
            return res.render('scouterSelect', {
                eventKey,
                eventName: eventKey,
                users: null,
                scheduleAvailable: false
            });
        }

        let schedule = {};
        let assignmentCounts = {};
        const MAX_MATCHES_PER_SCOUTER = 16;

        scouters.forEach(s => {
            schedule[s] = [];
            assignmentCounts[s] = 0;
        });

        matches.forEach(match => {
            const redTeams = match.alliances?.red?.team_keys || [];
            const blueTeams = match.alliances?.blue?.team_keys || [];
            const teams = [...redTeams, ...blueTeams];

            if (teams.length === 0) return;

            teams.forEach(team => {
                const eligibleScouters = scouters.filter(s => assignmentCounts[s] < MAX_MATCHES_PER_SCOUTER);

                if (eligibleScouters.length === 0) return;

                const minCount = Math.min(...eligibleScouters.map(s => assignmentCounts[s]));
                const candidates = eligibleScouters.filter(s => assignmentCounts[s] === minCount);

                const chosen = candidates[Math.floor(Math.random() * candidates.length)];
                schedule[chosen].push({
                    match: match.match_number || match.key,
                    team
                });

                assignmentCounts[chosen]++;
            });
        });

        scouters.forEach(s => usersData.users[s].shiftsAssigned = schedule[s].length);

        const schedulePath = path.join(__dirname, 'data', `schedule_${eventKey}.json`);
        fs.writeFileSync(schedulePath, JSON.stringify(schedule, null, 4));
        fs.writeFileSync(usersPath, JSON.stringify(usersData, null, 4));

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
        console.error('Error checking matches:', err.response?.data || err.message);
        res.send("Error checking matches.");
    }
});

app.post('/login', async (req, res) => {
    const { eventKey, username } = req.body;
    const schedulePath = path.join(__dirname, 'data', `schedule_${eventKey}.json`);
    const schedule = JSON.parse(fs.readFileSync(schedulePath));
    const assignedMatches = schedule[username] || [];

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
    const { eventKey, username } = data;

    const exportPath = path.join(__dirname, 'exports', `${eventKey}.xlsx`);
    const workbook = new ExcelJS.Workbook();

    if (fs.existsSync(exportPath)) await workbook.xlsx.readFile(exportPath);

    let sheet = workbook.getWorksheet("Summary");
    if (!sheet) {
        sheet = workbook.addWorksheet("Summary");
        sheet.addRow(Object.keys(data));
    }

    sheet.addRow(Object.values(data));
    await workbook.xlsx.writeFile(exportPath);

    const users = JSON.parse(fs.readFileSync(usersPath));
    if (Array.isArray(users.users)) {
        const u = users.users.find(u => u.username === username);
        if (u) u.shiftsSubmitted = (u.shiftsSubmitted || 0) + 1;
    } else {
        users.users[username].shiftsSubmitted = (users.users[username].shiftsSubmitted || 0) + 1;
    }
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

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));