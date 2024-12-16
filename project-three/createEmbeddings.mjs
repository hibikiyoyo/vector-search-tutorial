import { promises as fsp } from "fs";
import * as fs from "fs";
import { RecursiveCharacterTextSplitter } from "langchain/text_splitter";
import { MongoDBAtlasVectorSearch } from "langchain/vectorstores/mongodb_atlas";
import { OpenAIEmbeddings } from "langchain/embeddings/openai";
import { MongoClient } from "mongodb";
import "dotenv/config";
import * as path from "path";
import { ChunkDeeper} from "./chunkDeeper.mjs";
import iconv from "iconv-lite";

const client = new MongoClient(process.env.MONGODB_ATLAS_URI || "");
const dbName = "docs";
const collectionName = "vb6";
const collection = client.db(dbName).collection(collectionName);

// filter files with extensions
const extensions = /\.(txt|frm|cls|md|bas|vbp)$/i;
const docs_dir = "test";
let files = getAllFiles(docs_dir, extensions)

// const fileNames = await fsp.readdir(docs_dir);

for (const fileName of files) {
  const document = await fsp.readFile(`${fileName}`, "utf8");
  // console.log(document)
  const document_jp = iconv.decode(document, 'shift_jis');
  
  
  console.log(`Vectorizing ${fileName}`);

  console.log(`Vectorizing ${fileName}`);
  let extension = fileName.split('.').pop();
  let output = ''
  output = await makeDocumentForMd(document_jp)
  // if (extension === "md") {
  //   // output = await makeDocumentForMd(document_jp)
  // } else {
  //   output = makeDocument(document_jp)
  // }

  console.log(output)
  // continue;
  await MongoDBAtlasVectorSearch.fromDocuments(
    output,
    new OpenAIEmbeddings(),
    {
      collection,
      indexName: "default",
      textKey: "text",
      embeddingKey: "embedding",
    }
  );
}

console.log("Done: Closing Connection");
await client.close();

function getAllFiles(directory, extensions) {
  let files = [];
  const items =  fs.readdirSync(directory, { withFileTypes: true });
  items.forEach(item => {
      const itemPath = path.join(directory, item.name);
      if (item.isDirectory()) {
          files = files.concat(getAllFiles(itemPath, extensions)); // Recurse into subdirectory
      } else if (extensions.test(item.name)) {
          files.push(itemPath); // Add matching file
      }
  });

  return files;
}

function makeDocument(document) {
  const chunkDeeper = new ChunkDeeper()
  const output = chunkDeeper.createDocument(document, 3)
  return output
}

async function makeDocumentForMd(document) {
  // Using Spliter
  const splitter = RecursiveCharacterTextSplitter.fromLanguage("markdown", {
    chunkSize: 3000,
    chunkOverlap: 300,
  });

  const output = await splitter.createDocuments([document]);
  return output
}