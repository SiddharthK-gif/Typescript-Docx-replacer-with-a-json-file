# Typescript-Docx-replacer-with-a-json-file

**Download and unzip the file. then replace the docx file and the sample json file and run it according to this file.**


# How the file works:
The big picture: A .docx file is actually a ZIP archive containing XML files. The program unzips the document, finds your {{placeholders}} in the XML, swaps them with values from your JSON, then rezips it back into a valid .docx.

Step 1 — Read the template
typescriptconst zip = new PizZip(fs.readFileSync(absTemplate));
PizZip unzips the .docx into memory, exposing all the internal XML files (the document text lives in word/document.xml).

Step 2 — Read and flatten the JSON
typescriptconst data = flattenObject(rawData);
This is where dot-notation support comes from. A nested JSON like:
json{ "address": { "city": "San Francisco" } }
gets flattened into:
json{ "address": { "city": "SF" }, "address.city": "San Francisco" }
So when docxtemplater encounters {{address.city}} in the XML, it finds a matching key directly. The original nested structure is kept alongside so that loop blocks like {{#address}}{{city}}{{/address}} still work too.

Step 3 — Parse and render the template
typescriptconst doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
doc.render(data);
Docxtemplater scans through the XML looking for your delimiter tags ({{ and }}). It handles four kinds of tags:
TagBehaviour{{name}}Replaced with the value from JSON{{address.city}}Resolved via the flattened key{{#items}}…{{/items}}Loops — repeats the block for each item in the array{{#isRemote}}…{{/isRemote}}Conditional — renders only if the value is truthy
Crucially, it works at the XML level, not on plain text. This means all your formatting — bold, font size, colours, tables — is fully preserved in the output. It's not doing a simple find-and-replace on text; it's surgically editing the XML nodes.

Step 4 — Write the output
typescriptconst outBuffer = doc.getZip().generate({ type: "nodebuffer" });
fs.writeFileSync(outputPath, outBuffer);
The modified XML is rezipped back into a proper .docx file and written to disk.

Why placeholders sometimes break in Word
Word often splits a single {{name}} across multiple XML <w:r> (run) elements internally, especially if you edited the text after typing it. So the raw XML might look like:
xml<w:r><w:t>{{</w:t></w:r>
<w:r><w:t>name</w:t></w:r>
<w:r><w:t>}}</w:t></w:r>
Docxtemplater handles this automatically by merging split tags before processing, which is one of the main reasons to use it over a plain string replace.

# How to run:
The command was:
bashnode dist/index.js template.docx sample-data.json output.docx
Where:

template.docx → your input Word template
sample-data.json → your JSON data file
output.docx → the name you want for the filled result

To use your own files, just replace those with your own filenames. For example:
bashnode dist/index.js my-contract.docx my-data.json filled-contract.docx
The only requirement is that your .docx uses {{variable}} tags that match the keys in your JSON file. For example if your JSON has:
json{ "clientName": "Acme Corp" }
Then your Word document should have {{clientName}} wherever you want that value to appear.
