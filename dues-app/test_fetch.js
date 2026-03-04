const https = require('https');

function fetchGas() {
    return new Promise((resolve, reject) => {
        const url = 'https://script.google.com/macros/s/AKfycbyelhW9r6BVeHoS6ooGaV0Vah6tDPAMdCPmfPM6nGPfVcc2fv7dxQNDYh0wK2EmfyYFAQ/exec?action=getAllData';

        https.get(url, (res) => {
            if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
                // Handle redirect
                let redirUrl = res.headers.location;
                https.get(redirUrl, (redirRes) => {
                    let data = '';
                    redirRes.on('data', chunk => data += chunk);
                    redirRes.on('end', () => resolve(data));
                }).on('error', reject);
            } else {
                let data = '';
                res.on('data', chunk => data += chunk);
                res.on('end', () => resolve(data));
            }
        }).on('error', reject);
    });
}

fetchGas().then(data => {
    const parsed = JSON.parse(data);
    console.log("Total members found:", parsed.members ? parsed.members.length : 0);
    console.log("Total transactions found:", parsed.transactions ? parsed.transactions.length : 0);

    if (parsed.members) {
        console.log("First 10 members:", parsed.members.slice(0, 10).map(m => m.name));

        // Check for Kim Dong Yong
        const found = parsed.members.filter(m => String(m.name).replace(/\s/g, '').includes("김동용"));
        console.log("Search for 김동용 (no spaces):", found.length > 0 ? "Found names: " + found.map(m => m.name).join(', ') : "Not Found");

        if (found.length === 0) {
            console.log("All member names:", parsed.members.map(m => m.name).join(', '));
        }
    }
}).catch(console.error);
