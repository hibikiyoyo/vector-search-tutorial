import { promises as fsp } from "fs";
import * as fs from "fs";
import { RecursiveCharacterTextSplitter } from "langchain/text_splitter";
import { MongoDBAtlasVectorSearch } from "langchain/vectorstores/mongodb_atlas";
import { OpenAIEmbeddings } from "langchain/embeddings/openai";
import { MongoClient } from "mongodb";
import "dotenv/config";
import * as path from "path";
import {ParenthesesAwareTextSplitter} from "./chunk.mjs"
import { Document } from 'langchain/document'

const client = new MongoClient(process.env.MONGODB_ATLAS_URI || "");
const dbName = "docs";
const collectionName = "embeddings";
const collection = client.db(dbName).collection(collectionName);

// filter files with extensions
const extensions = /\.(txt|frm|cls)$/i;
const docs_dir = "_assets";
let files = getAllFiles(docs_dir, extensions)

// const fileNames = await fsp.readdir(docs_dir);

for (const fileName of files) {
  const document = await fsp.readFile(`${fileName}`, "utf8");
  // console.log(document)
  console.log(`Vectorizing ${fileName}`);
  

  // Spliter
  const splitter = RecursiveCharacterTextSplitter.fromLanguage("markdown", {
    chunkSize: 500,
    chunkOverlap: 50,
  });

  // Initialize the splitter with desired chunk size and overlap
  // const splitter = new ParenthesesAwareTextSplitter({ chunkSize: 1000, chunkOverlap: 50  });

  const output = await splitter.createDocuments([document]);
  console.log(output)

  // await MongoDBAtlasVectorSearch.fromDocuments(
  //   output,
  //   new OpenAIEmbeddings(),
  //   {
  //     collection,
  //     indexName: "default",
  //     textKey: "text",
  //     embeddingKey: "embedding",
  //   }
  // );
}

// console.log("Done: Closing Connection");
// await client.close();

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
