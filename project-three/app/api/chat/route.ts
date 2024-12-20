import { StreamingTextResponse, LangChainStream, Message } from 'ai';
import { ChatOpenAI } from 'langchain/chat_models/openai';
import { AIMessage, HumanMessage } from 'langchain/schema';

export const runtime = 'edge';

export async function POST(req: Request) {
  console.log("aaa")
  const { messages } = await req.json();
  const currentMessageContent = messages[messages.length - 1].content;

  const vectorSearch = await fetch("http://localhost:3000/api/vectorSearch", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: currentMessageContent,
  }).then((res) => res.json());
  console.log(vectorSearch)
  //If you are unsure and the answer is not explicitly written in the documentation, say "Sorry, I don't know how to help with that.
  const TEMPLATE = `You are an expert coder's assistant, helping to transform legacy VB6 form files into modern React.js applications. Please use all of your knowledge and any relevant documentation available on the internet to assist in this transformation. Given the following sections from the vbFile saved as vector in embedded mongodb answer the question using that information like the base knowledge, outputted in text or code. If you are not sure about the answer, then you can reseach in the internet for helping you to make the better answer.
  Context sections:
  ${JSON.stringify(vectorSearch)}

  Question: """
  ${currentMessageContent}
  """
  `;

  messages[messages.length -1].content = TEMPLATE;

  const { stream, handlers } = LangChainStream();

  const llm = new ChatOpenAI({
    modelName: "gpt-4o-mini",
    streaming: true,
  });

  llm
    .call(
      (messages as Message[]).map(m =>
        m.role == 'user'
          ? new HumanMessage(m.content)
          : new AIMessage(m.content),
      ),
      {},
      [handlers],
    )
    .catch(console.error);

  return new StreamingTextResponse(stream);
}
