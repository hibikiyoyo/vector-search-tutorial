import { promises as fsp } from "fs";
import * as fs from "fs";
import { RecursiveCharacterTextSplitter } from "langchain/text_splitter";
import { MongoDBAtlasVectorSearch } from "langchain/vectorstores/mongodb_atlas";
import { OpenAIEmbeddings } from "langchain/embeddings/openai";
import { MongoClient } from "mongodb";
import "dotenv/config";
import * as path from "path";
import {ParenthesesAwareTextSplitter} from "./chunk.mjs"
import { ChunkDeeper} from "./chunkDeeper.mjs"
import { Document } from 'langchain/document'

const client = new MongoClient(process.env.MONGODB_ATLAS_URI || "");
const dbName = "docs";
const collectionName = "embeddings";
const collection = client.db(dbName).collection(collectionName);

// filter files with extensions
const extensions = /\.(txt|frm|cls|md)$/i;
const docs_dir = "_assets";
let files = getAllFiles(docs_dir, extensions)

// const fileNames = await fsp.readdir(docs_dir);

for (const fileName of files) {
  const document = await fsp.readFile(`${fileName}`, "utf8");
  // console.log(document)
  
  console.log(`Vectorizing ${fileName}`);
  let extension = fileName.split('.').pop();
  let output = ''
  if (extension === "md") {
    output = await makeDocumentForMd(document)
  } else {
    output = makeDocument(document)
  }

  // console.log(output)

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
    chunkSize: 500,
    chunkOverlap: 50,
  });

  const output = await splitter.createDocuments([document]);
  return output
}