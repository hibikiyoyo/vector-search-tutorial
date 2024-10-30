import * as fs from "fs";
import { Document } from 'langchain/document'
let input = fs.readFileSync('_assets/test/frmRA_tree.txt', "utf8");

export class ChunkDeeper {
  createDocument(input, maxDepth) {

    try {
      // Parse the tree text into an AST
      const ast = parseTreeText(input);
      // Split the AST into chunks based on maximum depth
      const chunks = splitAST(ast, maxDepth);
      
      // Print the chunks
      // chunks.forEach((chunk, index) => {
      //   console.log(`Chunk ${index + 1}:\n${chunk}\n`);
      // });
      console.log("chunks")
      // Create documents
      const documents = chunks.
        filter((chunk, index) => {
          return !(chunk.trim() == '\\r\\n')
        })
        .map((chunk, index) => {
          return new Document({
          pageContent: chunk,
          metadata: { id: index + 1 },
          });
        });
      return documents
    } catch (error) {
      console.error('Error parsing tree text:', error.message);
    }
  }
}
// Helper function to parse the parenthetical tree text into an AST
function parseTreeText(text) {
    let index = 0;
  
    function skipWhitespace() {
      // console.log(text[index])
      while (/\s/.test(text[index]) || isWhitespace(text[index])) index++;
    }
    
    function isWhitespace(char) {
      return char === ' ' || char === '\t' || char === '\n' || char === '\r';
    }

    function parseNode() {
      skipWhitespace();
  
      if (text[index] !== '(') {
        throw new Error(`Expected '(', found '${text[index]}' at position ${index}`);
      }
  
      index++; // Skip '('
      skipWhitespace();
  
      // Parse the node label
      let label = '';
      while (text[index] && !isWhitespace(text[index]) && !/\s|\(|\)/.test(text[index])) {
        label += text[index++];
      }
  
      const node = { label, children: [] };
      skipWhitespace();
  
      // Parse children
      while (text[index] && text[index] !== ')') {
        if (text[index] === '(') {
          node.children.push(parseNode());
        } else {
          // Parse leaf node (token)
          let token = '';
          while (text[index] && !isWhitespace(text[index]) && !/\s|\(|\)/.test(text[index])) {
            token += text[index++];
          }
          node.children.push({ label: token, children: [] });
        }
        skipWhitespace();
      }
  
      if (text[index] !== ')') {
        throw new Error(`Expected ')', found '${text[index]}' at position ${index}`);
      }
  
      index++; // Skip ')'
      return node;
    }
  
    const ast = [];
    skipWhitespace();
  
    while (index < text.length) {
      if (text[index] === '(') {
        ast.push(parseNode());
      } else if (/\s/.test(text[index])) {
        skipWhitespace();
      } else if (isWhitespace(text[index])) {
        skipWhitespace();
      } else {
        throw new Error(`Unexpected character '${text[index]}' at position ${index}`);
      }
    }
  
    return ast;
  }
  
  // Function to serialize AST back to parenthetical notation
function serializeNode(node) {
    if (node.children.length === 0) {
      return node.label;
    }
    const childrenStr = node.children.map(serializeNode).join(' ');
    return `(${node.label} ${childrenStr})`;
  }
  
  // Function to split AST into chunks based on maximum depth
function splitAST(ast, maxDepth) {
    const chunks = [];
  
    function traverse(node, depth) {
      if (depth > maxDepth) {
        // Serialize the node and add to chunks
        const chunk = serializeNode(node);
        chunks.push(chunk);
        return;
      }
  
      // Traverse children
      node.children.forEach((child) => {
        traverse(child, depth + 1);
      });
    }
  
    ast.forEach((node) => {
      traverse(node, 1);
    });
  
    return chunks;
  }
  
  // Usage Example
  const treeText = `(startRule (module 
   (moduleBody (moduleBodyElement (subStmt (visibility Private)   Sub   (ambiguousIdentifier Command1_Click)   (argList ( )) 
      (block (blockStmt (letStmt (implicitCallStmt_InStmt (iCS_S_MembersCall (iCS_S_VariableOrProcedureCall (ambiguousIdentifier Text1)) (iCS_S_MemberCall . (iCS_S_VariableOrProcedureCall (ambiguousIdentifier (ambiguousKeyword Text))))))   =   (valueStmt (literal "Hello, world!"))))) 
   End Sub))) 
  ) <EOF>)`;

  // try {
  //   // Parse the tree text into an AST
  //   const ast = parseTreeText(input);
  
  //   // Split the AST into chunks based on maximum depth
  //   const maxDepth = 3; // Adjust as needed
  //   const chunks = splitAST(ast, maxDepth);
    
  //   // Print the chunks
  //   chunks.forEach((chunk, index) => {
  //     console.log(`Chunk ${index + 1}:\n${chunk}\n`);
  //   });

  //   // Create documents
  //   const documents = chunks.
  //     filter((chunk, index) => {
  //       return !(chunk.trim() == '\\r\\n')
  //     })
  //     .map((chunk, index) => {
  //       return new Document({
  //       pageContent: chunk,
  //       metadata: { id: index + 1 },
  //       });
  //     });
  //     console.log("hihi")
  // } catch (error) {
  //   console.error('Error parsing tree text:', error.message);
  // }