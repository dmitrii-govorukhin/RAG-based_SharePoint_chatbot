# SharePoint search chatbot (RAG-service)

**Project**: The chatbot lets users ask ordinary questions about information stored on the SharePoint site. It combines SharePoint Search with GPT-5.6 to produce concise answers grounded in information retrieved from SharePoint.

**Consumer**: Contra Costa County Employment and Human Services Department, CA

**Program languages**: Python, JavaScript

**Python libraries are used**: requests, json, uuid, datetime, fastapi, pydantic, requests_negotiate_sspi, docx, pypdf, openai

**Description**: 
The chatbot workflow:
1. The user enters a question in everyday language; no SharePoint search syntax is required.
2. JavaScript on SharePoint site sent the request to the backend server.
3. Backend server processes the request and sent it to ChatGPT through OpenAI API.
4. The request is converted into a concise search phrase intended to locate the most relevant SharePoint information.
5. SharePoint Search finds matching documents and returns their titles, links, file types, relevance ranks, and short search snippets.
6. The relevant SharePoint results and snippets become the source material used to answer the user's questions.
7. GPT is instructed to answer only from the supplied SharePoint information. If the results are insufficient, it should say that the information was not found.
8. The chatbot displays the documents found by SharePoint, with links, ranks, and snippets, allowing users to open and review the source material themselves.

The project file structure:
 - backend
    - main.py - backend script
 - frontend
    - chatbot.html - frontend script and chatbot web interface

**Results / Key Findings:** The chatbot connects artificial intelligence to the organization’s knowledge base to provide accurate answers based on up-to-date company documents, rather than just the neural network’s collective memory. 

**Illustrations**: Step-by-step process, Chatbot interface examples.

![alt text](https://github.com/dmitrii-govorukhin/SharePoint-search-chatbot-RAG-service-/blob/main/img/process.png?raw=true)

![alt text](https://github.com/dmitrii-govorukhin/SharePoint-search-chatbot-RAG-service-/blob/main/img/example_1.png?raw=true)

![alt text](https://github.com/dmitrii-govorukhin/SharePoint-search-chatbot-RAG-service-/blob/main/img/example_2.png?raw=true)

