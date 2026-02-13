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

const DEV_MODE = true;
const DEV_EVENT_KEY = "devtest";

// HOME
app.get('/', async (req, res) => {
    if (DEV_MODE) {
        return res.render('index', {
            events: [{ key: DEV_EVENT_KEY, name: "DEV TEST EVENT", city: "Test City" }]
        });
    }

    try {
        const currentYear = new Date().getFullYear();
        const response = await axios.get(
            `https://www.thebluealliance.com/api/v3/team/frc2068/events/${currentYear}`,
            { headers: { 'X-TBA-Auth-Key': TBA_KEY } }
        );

        res.render('index', { events: response.data });
    } catch (err) {
        console.error(err.response?.data || err.message);
        res.send("Error loading Team 2068 events.");
    }
});

// EVENT SELECTED
app.post('/select-event', async (req, res) => {
    const { eventKey } = req.body;

    if (DEV_MODE && eventKey === DEV_EVENT_KEY) {
        const matches = require('./data/dev_matches.json');
        const usersData = JSON.parse(fs.readFileSync(usersPath));
        const scouters = Object.keys(usersData.users).filter(u => usersData.users[u].scouter);

        if (scouters.length < 6) return res.send("Need at least 6 scouters for full coverage.");

        const shuffle = arr => arr.sort(() => Math.random() - 0.5);
        let schedule = {};
        scouters.forEach(s => schedule[s] = []);

        matches.forEach(match => {
            const teams = [...match.alliances.red.teams, ...match.alliances.blue.teams];
            const scouterRotation = shuffle([...scouters]);

            teams.forEach((team, i) => {
                const assignedScouter = scouterRotation[i % scouterRotation.length];
                schedule[assignedScouter].push({ match: match.match_number, team });
            });
        });

        scouters.forEach(s => usersData.users[s].shiftsAssigned = schedule[s].length);

        const schedulePath = path.join(__dirname, 'data', `schedule_${eventKey}.json`);
        fs.writeFileSync(schedulePath, JSON.stringify(schedule, null, 4));
        fs.writeFileSync(usersPath, JSON.stringify(usersData, null, 4));

        return res.render('scouterSelect', {
            eventKey,
            users: usersData.users,
            scheduleAvailable: true
        });
    }

    // ===== REAL MODE =====
    try {
        const matchesRes = await axios.get(
            `https://www.thebluealliance.com/api/v3/event/${eventKey}/matches/simple`,
            { headers: { 'X-TBA-Auth-Key': TBA_KEY } }
        );

        const matches = matchesRes.data.filter(m => m.comp_level === "qm");

        if (!matches || matches.length === 0) {
            return res.render('scouterSelect', {
                eventKey,
                users: null,
                scheduleAvailable: false
            });
        }

        const usersData = JSON.parse(fs.readFileSync(usersPath));
        const scouters = Object.keys(usersData.users).filter(u => usersData.users[u].scouter);

        if (scouters.length < 6) return res.send("Need at least 6 scouters for full coverage.");

        const shuffle = arr => arr.sort(() => Math.random() - 0.5);
        let schedule = {};
        scouters.forEach(s => schedule[s] = []);

        matches.forEach(match => {
            const teams = [...match.alliances.red.teams, ...match.alliances.blue.teams];
            const scouterRotation = shuffle([...scouters]);

            teams.forEach((team, i) => {
                const assignedScouter = scouterRotation[i % scouterRotation.length];
                schedule[assignedScouter].push({ match: match.match_number, team });
            });
        });

        scouters.forEach(s => usersData.users[s].shiftsAssigned = schedule[s].length);

        const schedulePath = path.join(__dirname, 'data', `schedule_${eventKey}.json`);
        fs.writeFileSync(schedulePath, JSON.stringify(schedule, null, 4));
        fs.writeFileSync(usersPath, JSON.stringify(usersData, null, 4));

        res.render('scouterSelect', {
            eventKey,
            users: usersData.users,
            scheduleAvailable: true
        });

    } catch (err) {
        console.log(err.response?.data || err.message);
        res.send("Error checking matches.");
    }
});

// LOGIN
app.post('/login', (req, res) => {
    const { eventKey, username } = req.body;
    const schedulePath = path.join(__dirname, 'data', `schedule_${eventKey}.json`);
    const schedule = JSON.parse(fs.readFileSync(schedulePath));
    const assignedMatches = schedule[username] || [];

    res.render('dashboard', {
        eventKey,
        username,
        matches: assignedMatches
    });
});

// SCOUT MATCH
app.post('/scout', (req, res) => {
    const { eventKey, username, matchNumber, teamNumber } = req.body;

    res.render('scoutForm', {
        eventKey,
        username,
        matchNumber,
        teamNumber
    });
});

// SUBMIT SCOUT
app.post('/submit-scout', async (req, res) => {
    const data = req.body;
    const { eventKey, username } = data;

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

    const users = JSON.parse(fs.readFileSync(usersPath));
    users.users[username].shiftsSubmitted++;
    fs.writeFileSync(usersPath, JSON.stringify(users, null, 4));

    res.send("<h2>Submitted Successfully</h2><a href='/'>Home</a>");
});

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
