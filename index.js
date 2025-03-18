import fetch from 'node-fetch';
import * as cheerio from 'cheerio';
import fs from 'fs';
import ExcelJS from 'exceljs';

let result = [];

async function scrapeThegef() {
    let response = await fetch("https://www.thegef.org/who-we-are/focal-points", {
        headers: {
            "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "accept-language": "en-US,en;q=0.9",
            "cache-control": "max-age=0",
            "sec-ch-ua": "\"Chromium\";v=\"134\", \"Not:A-Brand\";v=\"24\", \"Google Chrome\";v=\"134\"",
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "\"Windows\"",
            "sec-fetch-dest": "document",
            "sec-fetch-mode": "navigate",
            "sec-fetch-site": "none",
            "sec-fetch-user": "?1",
            "upgrade-insecure-requests": "1"
        },
        method: "GET"
    });

    if (!response.ok) {
        console.error(`Failed to fetch page: ${response.status} ${response.statusText}`);
        return;
    }

    let html = await response.text();
    let $ = cheerio.load(html);

    $(".focal-point-country").each((index, element) => {
        let country = $(element).find("h2").text().trim();

        let persons = $(element).html().split(/<h4>/).slice(1);

        persons.forEach(personHtml => {
            let person$ = cheerio.load("<h4>" + personHtml); 
            let name = person$("h4").text().trim();

            let parts = personHtml.split("<br>");
            let title = parts.length > 1 ? parts[1].trim() : "";
            let organization = parts.length > 2 ? parts[2].trim() : "";

            let emailRegex = /([\w.-]+)\s*(?:\[a t \]|\(A T\)| at |\(AT\)|\(a t \)|\(at\)|\(AT\))\s*([\w.-]+)/gi;
            let emails = [];
            let match;
            while ((match = emailRegex.exec(personHtml)) !== null) {
                emails.push(`${match[1]}@${match[2]}`);
            }

            let email1 = emails.length > 0 ? emails[0] : "";
            let email2 = emails.length > 1 ? emails[1] : "";

            result.push({ country, name, title, organization, email1, email2 });
        });
    });
}

async function scrapeUnep() {
    let response = await fetch("https://www.unep.org/inc-plastic-pollution/national-focal-points", {
        "headers": {
            "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "accept-language": "en-US,en;q=0.9",
            "cache-control": "max-age=0",
            "priority": "u=0, i",
            "sec-ch-ua": "\"Chromium\";v=\"134\", \"Not:A-Brand\";v=\"24\", \"Google Chrome\";v=\"134\"",
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "\"Windows\"",
            "sec-fetch-dest": "document",
            "sec-fetch-mode": "navigate",
            "sec-fetch-site": "none",
            "sec-fetch-user": "?1",
            "upgrade-insecure-requests": "1"
        },
        "referrerPolicy": "strict-origin-when-cross-origin",
        "body": null,
        "method": "GET",
        "mode": "cors",
        "credentials": "include"
    });

    if (!response.ok) {
        console.error(`Failed to fetch page: ${response.status} ${response.statusText}`);
        return;
    }

    let html = await response.text();
    let $ = cheerio.load(html);

    $("table.cols-2 tbody tr").each((index, element) => {
        let country = $(element).closest("table").find(".inc_country").text().trim();
        let name = $(element).find(".inc_title").text().trim();

        let title = "";
        let organization = "";
        let phone = "";
        let email1 = "";
        let email2 = "";

        $(element).find(".inc_body p").each((_, p) => {
            let text = $(p).text().trim();

            if (text.startsWith("Position:")) {
                title = text.replace("Position:", "").trim();
            } else if (text.startsWith("Office:")) {
                organization = text.replace("Office:", "").trim();
            } else if (text.startsWith("Phone number:")) {
                phone = text.replace("Phone number:", "").trim();
            } else if (text.startsWith("Email:")) {
                let emails = $(p).find("a").map((_, a) => $(a).text().trim()).get();
                email1 = emails.length > 0 ? emails[0] : "";
                email2 = emails.length > 1 ? emails[1] : "";
            }
        });

        result.push({ country, name, title, organization, email1, email2 });
    });
}

async function scrapeBrsmeas() {
    let response = await fetch("https://informea.pops.int/Contacts2/brsContacts.svc/Contacts?$callback=jQuery112407286579056920082_1742257663983&%24inlinecount=allpages&%24format=json&%24top=1500&%24orderby=brsPartyNameEn%2CbrsTreatyName%2Cprimary+desc%2ClastName", {
        "headers": {
            "accept": "*/*",
            "accept-language": "en-US,en;q=0.9",
            "sec-ch-ua": "\"Chromium\";v=\"134\", \"Not:A-Brand\";v=\"24\", \"Google Chrome\";v=\"134\"",
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "\"Windows\"",
            "sec-fetch-dest": "script",
            "sec-fetch-mode": "no-cors",
            "sec-fetch-site": "cross-site",
            "sec-fetch-storage-access": "active",
            "Referer": "https://www.brsmeas.org/",
            "Referrer-Policy": "strict-origin-when-cross-origin"
        },
        "body": null,
        "method": "GET"
    });

    if (!response.ok) {
        console.error(`Failed to fetch page: ${response.status} ${response.statusText}`);
        return;
    }

    const text = await response.text();

    const jsonText = text.match(/\{.*\}/s)?.[0];
    const json = JSON.parse(jsonText);

    json.d.results.forEach(item => {
        const country = item.brsPartyNameEn || item.country || "";
        const name = item.brsFullName || `${item.prefix || ""} ${item.firstName || ""} ${item.lastName || ""}`.trim();
        const title = item.position || "";
        const organization = item.institution || "";
        const emails = item.brsEmails ? item.brsEmails.split(/[,;]/).map(email => email.trim()) : [];
        const email1 = emails[0] || "";
        const email2 = emails[1] || "";

        result.push({ country, name, title, organization, email1, email2 });
    });
}

async function exportExcel() {
    let workbook = new ExcelJS.Workbook();
    let worksheet = workbook.addWorksheet("Scraped Data");

    worksheet.columns = [
        { header: "Country", key: "country", width: 20 },
        { header: "Name", key: "name", width: 30 },
        { header: "Title", key: "title", width: 40 },
        { header: "Organization", key: "organization", width: 50 },
        { header: "Email 1", key: "email1", width: 30 },
        { header: "Email 2", key: "email2", width: 30 }
    ];

    result.forEach(item => worksheet.addRow(item));

    await workbook.xlsx.writeFile("scraped_data.xlsx");
    console.log("Data exported to scraped_data.xlsx");
}

async function main() {
    await scrapeThegef();
    await scrapeUnep();
    await scrapeBrsmeas();

    if (result.length === 0) {
        console.log("No data found. Check your selectors.");
        return;
    }

    let seenNames = new Set();
    result = result.filter(entry => {
        if (seenNames.has(entry.name)) {
            return false; 
        }
        seenNames.add(entry.name);
        return true;
    });

    result.sort((a, b) => a.country.localeCompare(b.country));
    exportExcel();
}

main();