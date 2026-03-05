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
            return res.render('scouterSelect', {
                eventKey,
                eventName: eventKey,
                users: usersData.users,
                scheduleAvailable: true
            });
        }

        const scouters = Object.keys(usersData.users)
            .filter(u => usersData.users[u].scouter);

        if (scouters.length < 1) {
            return res.send("Need at least 1 scouter.");
        }

        const matchesRes = await axios.get(
            `https://www.thebluealliance.com/api/v3/event/${eventKey}/matches`,
            { headers: { 'X-TBA-Auth-Key': TBA_KEY } }
        );

        const matches = matchesRes.data
            .filter(m => m.comp_level === "qm")
            .sort((a, b) => a.match_number - b.match_number);

        let schedule = {};
        let counts = {};
        const MAX_MATCHES = 16;

        scouters.forEach(s => {
            schedule[s] = [];
            counts[s] = 0;
        });

        matches.forEach(match => {
            const teams = [
                ...(match.alliances?.red?.team_keys || []),
                ...(match.alliances?.blue?.team_keys || [])
            ];

            teams.forEach(team => {
                const eligible = scouters.filter(s => counts[s] < MAX_MATCHES);
                if (!eligible.length) return;

                const min = Math.min(...eligible.map(s => counts[s]));
                const candidates = eligible.filter(s => counts[s] === min);
                const chosen = candidates[Math.floor(Math.random() * candidates.length)];

                schedule[chosen].push({
                    match: match.match_number,
                    team
                });

                counts[chosen]++;
            });
        });

        fs.writeFileSync(schedulePath, JSON.stringify(schedule, null, 4));

        res.render('scouterSelect', {
            eventKey,
            eventName: eventKey,
            users: usersData.users,
            scheduleAvailable: true
        });

    } catch (err) {
        console.error(err.response?.data || err.message);
        res.send("Error generating schedule.");
    }
});

app.post('/login', (req, res) => {
    const { eventKey, username } = req.body;
    const schedulePath = path.join(__dirname, 'data', `schedule_${eventKey}.json`);

    if (!fs.existsSync(schedulePath)) {
        return res.send("Schedule not generated yet.");
    }

    const schedule = JSON.parse(fs.readFileSync(schedulePath));
    const assignedMatches = schedule[username] || [];

    res.render('dashboard', {
        eventKey,
        eventName: eventKey,
        username,
        matches: assignedMatches
    });
});

app.post('/scout', (req, res) => {
    const { eventKey, username, matchNumber, teamNumber } = req.body;
    res.render('scoutForm', { eventKey, username, matchNumber, teamNumber });
});


app.post('/submit-scout', async (req, res) => {
    try {
        const data = req.body;
        const { eventKey, username, matchNumber, teamNumber } = data;

        // Ensure public exports folder exists
        const exportsDir = path.join(__dirname, 'public', 'exports');
        if (!fs.existsSync(exportsDir)) {
            fs.mkdirSync(exportsDir, { recursive: true });
        }

        const exportPath = path.join(exportsDir, `${eventKey}.xlsx`);
        const workbook = new ExcelJS.Workbook();

        if (fs.existsSync(exportPath)) {
            await workbook.xlsx.readFile(exportPath);
        }

        let sheet = workbook.getWorksheet("Summary");

        if (!sheet) {
            sheet = workbook.addWorksheet("Summary");
            sheet.columns = Object.keys(data).map(key => ({
                header: key,
                key: key
            }));
        }

        sheet.addRow(data);
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

        res.render('submitSuccess', {
            username,
            eventName: eventKey,
            fileLink: `/exports/${eventKey}.xlsx`
        });

    } catch (err) {
        console.error("Submit scout error:", err);
        res.send("Error saving scout data.");
    }
});

app.get('/download/:eventKey', (req, res) => {
    const filePath = path.join(__dirname, 'public', 'exports', `${req.params.eventKey}.xlsx`);

    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.send("File not found.");
    }
});


app.listen(PORT, () =>
    console.log(`Server running on http://localhost:${PORT}`)
);