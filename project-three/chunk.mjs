// Import the TextSplitter base class
// const { TextSplitter } = require('langchain/text_splitter');
import { TextSplitter } from "langchain/text_splitter";
export class ParenthesesAwareTextSplitter extends TextSplitter {
  constructor({ chunkSize = 1000, chunkOverlap = 50 } = {}) {
    super();
    this.chunkSize = chunkSize; // Maximum number of characters per chunk
    this.chunkOverlap = chunkOverlap; // Overlap between chunks (if needed)
  }

  splitText(text) {
    const tokens = this._tokenizeText(text);
    const chunks = [];
    let currentChunk = [];
    let currentLength = 0;
    let balance = 0; // Parentheses balance counter

    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i];
      currentChunk.push(token);
      currentLength += token.length;

      if (token === '(') {
        balance += 1;
      } else if (token === ')') {
        balance -= 1;
      }

      // Check if current chunk should be split
      if (currentLength >= this.chunkSize && balance === 0) {
        chunks.push(currentChunk.join('').trim());
        // Handle chunk overlap if needed
        if (this.chunkOverlap > 0) {
          currentChunk = currentChunk.slice(-this.chunkOverlap);
          currentLength = currentChunk.reduce((acc, val) => acc + val.length, 0);
        } else {
          currentChunk = [];
          currentLength = 0;
        }
      }
    }

    // Add any remaining tokens as the last chunk
    if (currentChunk.length > 0) {
      chunks.push(currentChunk.join('').trim());
    }

    return chunks;
  }

  _tokenizeText(text) {
    // Regular expression to split the text into tokens
    const tokenPattern = /(\(|\)|\s+|[^\s()]+)/g;
    const tokens = text.match(tokenPattern);
    return tokens || [];
  }
}


export class DepthBasedTextSplitter extends TextSplitter {
    constructor({ maxDepth = 3 } = {}) {
      super();
      this.maxDepth = maxDepth;
    }
  
    splitText(text) {
      const tokens = this._tokenizeText(text);
      const chunks = [];
      let currentChunk = [];
      let currentDepth = 0;
  
      for (let token of tokens) {
        currentChunk.push(token);
  
        if (token === '(') {
          currentDepth += 1;
        } else if (token === ')') {
          currentDepth -= 1;
        }
  
        // Split when we return to the desired depth
        if (currentDepth <= this.maxDepth && token === ')') {
          chunks.push(currentChunk.join('').trim());
          currentChunk = [];
        }
      }
  
      // Add any remaining tokens
      if (currentChunk.length > 0) {
        chunks.push(currentChunk.join('').trim());
      }
  
      return chunks;
    }
  
    _tokenizeText(text) {
      const tokenPattern = /(\(|\)|\s+|[^\s()]+)/g;
      const tokens = text.match(tokenPattern);
      return tokens || [];
    }
  }

// Usage Example
// const treeText = `(startRule (module 
//  (moduleBody (moduleBodyElement (subStmt (visibility Private)   Sub   (ambiguousIdentifier Command1_Click)   (argList ( )) 
//     (block (blockStmt (letStmt (implicitCallStmt_InStmt (iCS_S_MembersCall (iCS_S_VariableOrProcedureCall (ambiguousIdentifier Text1)) (iCS_S_MemberCall . (iCS_S_VariableOrProcedureCall (ambiguousIdentifier (ambiguousKeyword Text))))))   =   (valueStmt (literal "Hello, world!"))))) 
//  End Sub))) 
// ) <EOF>)`;

// // Initialize the splitter with desired chunk size and overlap
// const splitter = new ParenthesesAwareTextSplitter({ chunkSize: 200 });
// // Split the tree text
// const chunks = splitter.splitText(treeText);

// // Print the chunks
// chunks.forEach((chunk, index) => {
//   console.log(`Chunk ${index + 1}:\n${chunk}\n`);
// });

// console.log("tesst")

// const depthSplitter = new DepthBasedTextSplitter({ maxDepth: 2 });
// const depthChunks = depthSplitter.splitText(treeText);
// depthChunks.forEach((chunk, index) => {
//     console.log(`Chunk ${index + 1}:\n${chunk}\n`);
//   });
