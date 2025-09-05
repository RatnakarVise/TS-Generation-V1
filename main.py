import os
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse
from app.agents.content_writer_agent import ContentWriterAgent
from app.doc.doc_constructor_agent import build_document
from app.doc.flow_diagram_agent import FlowDiagramAgent

app = FastAPI(title="ABAP Technical Spec Generator")

# @app.get("/health")
# async def health():
#     return {"status": "ok"}

@app.post("/generate_doc")
async def generate_doc(payload: dict):
    try:
        # 1. Create ContentWriter agent and generate section contents
        writer_agent = ContentWriterAgent()
        results = writer_agent.run(payload)
        # results: [{'section_name': ..., 'content': ...}, ...]
        
        # 2. Get the doc section template from the agent's own template
        sections = []
        for sec in writer_agent.template_sections:
            # Guess section type for flow diagram, table, else text:
            sec_type = "text"
            sec_title = sec["title"]
            if "diagram" in sec_title.lower():
                sec_type = "diagram"
            elif "| " in sec["content"]:  # crude: maybe sample is a table in the template
                sec_type = "table"
            sections.append({
                "title": sec_title,
                "type": sec_type
            })

        # 3. Build the document using response list, template structure, and diagram agent
        diagram_agent = FlowDiagramAgent()
        doc = build_document(results, sections, flow_diagram_agent=diagram_agent, diagram_dir="diagrams")

        # 4. Save to file
        output_filename = "Technical Specification Document.docx"
        output_path = os.path.abspath(output_filename)
        doc.save(output_path)

        # 5. Return the .docx file as HTTP file download
        return FileResponse(output_path,
                            filename=output_filename,
                            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": f"Failed to generate document: {str(e)}"})