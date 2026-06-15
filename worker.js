import * as XLSX from "xlsx";

export default {
    async fetch(request, env) {

        if (request.method !== "POST") {
            return new Response("OK");
        }

        const update = await request.json();

        const document = update?.message?.document;

        if (!document) {
            return new Response("No file");
        }

        try {

            const fileId = document.file_id;

            const fileInfo = await fetch(
                `https://api.telegram.org/bot${env.BOT_TOKEN}/getFile?file_id=${fileId}`
            ).then(r => r.json());

            const filePath = fileInfo.result.file_path;

            const userFile = await fetch(
                `https://api.telegram.org/file/bot${env.BOT_TOKEN}/${filePath}`
            );

            const userBuffer = await userFile.arrayBuffer();

            const templateFile = await fetch(
                "https://raw.githubusercontent.com/SeinUser/excel-bot/main/%D0%94%D0%BE%D0%BC%D0%B0PN.xlsx"
            );

            const templateBuffer = await templateFile.arrayBuffer();

            const resultFile = processExcel(
                userBuffer,
                templateBuffer
            );

            const chatId = update.message.chat.id;

            await sendDocument(
                env.BOT_TOKEN,
                chatId,
                resultFile
            );

        } catch (err) {

            await sendMessage(
                env.BOT_TOKEN,
                update.message.chat.id,
                "Ошибка обработки файла"
            );

            console.error(err);
        }

        return new Response("OK");
    }
};

function processExcel(userFileBuffer, templateBuffer) {

    const userWorkbook = XLSX.read(userFileBuffer);

    const userSheet =
        userWorkbook.Sheets[userWorkbook.SheetNames[0]];

    const templateWorkbook =
        XLSX.read(templateBuffer);

    const templateSheet =
        templateWorkbook.Sheets[templateWorkbook.SheetNames[0]];

    const counts12510 = {};
    const counts12511 = {};

    const range =
        XLSX.utils.decode_range(userSheet["!ref"]);

    for (let row = 20; row <= range.e.r + 1; row++) {

        const position =
            String(userSheet[`E${row}`]?.v ?? "").trim();

        if (!position) continue;

        const prefix =
            String(userSheet[`K${row}`]?.v ?? "")
                .substring(0, 5);

        if (prefix === "12510") {
            counts12510[position] =
                (counts12510[position] || 0) + 1;
        }

        if (prefix === "12511") {
            counts12511[position] =
                (counts12511[position] || 0) + 1;
        }
    }

    for (let row = 4; row <= 13; row++) {

        const position =
            String(templateSheet[`C${row}`]?.v ?? "").trim();

        templateSheet[`D${row}`] = {
            t: "n",
            v: counts12510[position] || 0
        };

        templateSheet[`E${row}`] = {
            t: "n",
            v: counts12511[position] || 0
        };
    }

    return XLSX.write(templateWorkbook, {
        type: "buffer",
        bookType: "xlsx"
    });
}

async function sendMessage(token, chatId, text) {

    await fetch(
        `https://api.telegram.org/bot${token}/sendMessage`,
        {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                chat_id: chatId,
                text
            })
        }
    );
}

async function sendDocument(token, chatId, fileBuffer) {

    const formData = new FormData();

    formData.append("chat_id", chatId);

    formData.append(
        "document",
        new Blob([fileBuffer]),
        "ДомаPN.xlsx"
    );

    await fetch(
        `https://api.telegram.org/bot${token}/sendDocument`,
        {
            method: "POST",
            body: formData
        }
    );
}
