import puppeteer from "puppeteer"
import xlsx from "node-xlsx"
import { writeFileSync } from "fs"

const MAIN_URL = "https://exporthungary.gov.hu/kulgazdasagi-attasek"

const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms))

function parsePeople(text) {
    // remove first line
    text = text.split("\n").slice(1).join("\n")

    // if multiple empty lines, only leave one empty line
    text = text.replace(/\n\n+/g, "\n")

    const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    const people = [];
    let current = {};

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        // Új név, ha nincs benne ':' és nem telefonszám
        if (!line.includes(':') && !line.toLowerCase().includes('telefonszám')) {
            if (line.toLowerCase().includes('tevékenys')) {
                current.note = line;
                continue;
            }
            if (Object.keys(current).length > 0) {
                people.push(current);
                current = {};
            }
            current.name = line;
        } else if (line.startsWith('Cím:')) {
            current.address = line.replace('Cím:', '').trim();
        } else if (line.startsWith('Email:')) {
            // Eltávolítjuk az esetleges kettőspontokat és szóközöket
            current.email = line.replace('Email:', '').replace(':', '').trim();
        } else if (line.toLowerCase().startsWith('hivatali telefonszám')) {
            current.workTel = line.replace(/Hivatali telefonszám:/i, '').trim();
        } else if (line.toLowerCase().startsWith('mobil telefonszám')) {
            current.homeTel = line.replace(/Mobil telefonszám:/i, '').replace(/Mobil Telefonszám/i, '').replace('+ ', '+').trim();
        }
    }
    // Az utolsó embert is hozzáadjuk, ha van adat
    if (Object.keys(current).length > 0) {
        people.push(current);
    }
    return people;
}

const main = async () => {
    const browser = await puppeteer.launch({
        headless: true,
        defaultViewport: null
    });

    // get page
    const page = (await browser.pages())[0]

    await page.setRequestInterception(true);

    page.on('request', (request) => {
        if (request.resourceType() === 'image') {
            request.abort(); // Képek letiltása
        } else {
            request.continue();
        }
    });

    await page.goto(MAIN_URL)

    const contentSelector = await page.$("#contentselector")

    const options = await contentSelector.$$("option")

    const optionsArray = await Promise.all(options.map(async option => await option.evaluate(el => [el.textContent, el.value])))

    const allPeople = []

    let i = 0

    for (const option of optionsArray) {
        await page.goto(`${MAIN_URL}?${option[1]}`)

        const content = await page.$("#contentselector-content")

        const contentText = await content.evaluate(el => el.textContent)

        const people = parsePeople(contentText)

        allPeople.push(...people.map(p => ({ ...p, country: option[0] })))

        i++
        console.log(`${i}/${optionsArray.length}`)
    }

    console.log("Creating excel table...")

    // create an excel table
    /// map allPeople to an array of objects with the keys name, address, email, workTel, homeTel, country
    const excelData = [
        ["Név", "Ország", "Cím", "Email", "Hivatali telefonszám", "Mobil telefonszám", "Megjegyzés"],
        ...allPeople.map(p => ([
            p.name,
            p.country,
            p.address,
            p.email,
            p.workTel,
            p.homeTel,
            p.note
        ]))]

    // create an excel table
    const buffer = xlsx.build([{ name: "People", data: excelData }])

    // write to file
    writeFileSync(`./${new Date().toISOString().split('T')[0]}.xlsx`, buffer)

    console.log("Excel table created")
    console.log("Closing browser...")
    await browser.close()
}

main()