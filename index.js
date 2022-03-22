const express = require('express');
const { MongoClient } = require('mongodb');
const { Document, Packer, Paragraph, TextRun } = require("docx");
const cors = require('cors');
require('dotenv').config()
const app = express();
const WordExtractor = require("word-extractor");
const fsPromises = require('fs').promises;
const readFileName = 'template.doc';
const writeFileName = 'proposal.doc';
const fs = require('fs');
const port = process.env.PORT || 5000;

// middleware
app.use(cors());
app.use(express.json());


const uri = `mongodb+srv://${process.env.DB_USER}:${process.env.DB_PASS}@cluster0.poyqe.mongodb.net/myFirstDatabase?retryWrites=true&w=majority`;
// console.log(uri)

const client = new MongoClient(uri);

async function run() {
    try {

        await client.connect();
        console.log('Connected to database');

        const database = client.db('customer-relation-management');
        const usersCollection = database.collection('users');

        // GET API
        app.post('/downloadPersonal', async (req, res) => {
            const loginData = req.body;
            const { projectBased, projectTime } = loginData;
            // console.log(projectBased,projectTime,loginData)
            const extractor = new WordExtractor();
            const extracted = extractor.extract(readFileName);
            let text = ''
            await extracted.then(function (doc) { text = doc.getBody(); });
            // console.log(text)
            // replace the things which needed
            let replacedName = await text.replace(/projectName/g, projectBased)
            let replacedMonths = await replacedName.replace(/projectMonth/g, projectTime)
            // split the array for difference the heder and text
            const myArray = await replacedMonths.split("\n\n\n");
            const userHeder = await myArray[0]
            const userText = await myArray[1]
            // make  sting for file
            const doc = new Document({
                sections: [{
                    properties: {},
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: userHeder,
                                    bold: true,
                                    size: 50,
                                    color: '#00B050',
                                }),

                            ],
                        }),
                        new Paragraph({}),
                        new Paragraph({}),
                        new Paragraph({}),
                        new Paragraph({
                            children: [
                                new TextRun(userText),

                            ],
                        }),
                    ],
                }],
            });
            // now write the string on file
            await Packer.toBuffer(doc).then((buffer) => {
                fs.writeFileSync(writeFileName, buffer);
            });

            res.download('./proposal.doc')
        })
        // POST API 
        // insert one
        app.post('/user', async (req, res) => {
            const loginData = req.body;
            const result = await usersCollection.findOne({ email: loginData.email });
            if (result.password === loginData.password) {
                res.json({ message: 'success' })
            } else {
                res.json({ message: 'failed' })

            }
        })
        app.post('/adduser', async (req, res) => {
            const loginData = req.body;
            let result = await usersCollection.findOne({ email: loginData.email });
            if (result) {
                res.json({ message: 'already have account' })
            } else {

                result = await usersCollection.insertOne(loginData);
                res.json(result)
            }
        })




    } finally {
        // await client.close();
    }
}

run().catch(console.dir);




app.listen(port, () => {
    console.log('your node server is running ', port)

})