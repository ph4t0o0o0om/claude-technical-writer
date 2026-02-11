from claude_agent_sdk import (
    ClaudeSDKClient, 
    SystemMessage, 
    ThinkingBlock, 
    ToolUseBlock, 
    create_sdk_mcp_server, 
    ClaudeAgentOptions, 
    AssistantMessage, 
    TextBlock, 
    tool, 
    AgentDefinition
)
from loguru import logger
from typing import Any
from pydantic import BaseModel
import asyncio
from dotenv import load_dotenv
import os

load_dotenv()


APPROVED_TOOLS=[
    "Read", 
    "Grep", 
    "Bash", 
    "KillShell", 
    "BashOutput", 
    "Fetch", 
    "WebSearch", 
    "ExitPlanMode", 
    "SlashCommand", 
    "WebFetch",
    "Task", 
    "Glob", 
    "Grep", 
    "TodoWrite",
    "Skill",
    "MultiEdit"
]

research_agent = AgentDefinition(
    description="Research and Analysis Agent that walks through the OWASP Juice Shop application like a regular user and captures information about the platform, as well as relevant screenshots",
    prompt="""
    You are a research and analysis agent that walks through the OWASP Juice Shop application like a regular user and captures information about the platform, as well as relevant screenshots.
    You need to use the Agent Browser skill to be able to access the platform, log in, and access all the functionality related to the platform in the demo server of OWASP Juice Shop. 
    You need to capture relevant screenshots at strategic junctures and use these screenshots in the guide embedded in the text sections, to ensure flow of the content with the right images (like a textbook), which will be done by the technical writer agent. 
    You need to capture information about the platform, including the features, capabilities, and workflows of the platform.
    You should not ever delete any object in the demo server of the platform.
    """,
)

technical_writer_agent = AgentDefinition(
    description="Technical Writer Agent that writes a user manual for the OWASP Juice Shop application",
    prompt="""
    You are a technical writer agent that writes a user manual for the OWASP Juice Shop application.
    You need to use the information captured by the research agent to write a user manual for the OWASP Juice Shop application.
    You need to write a user manual that is clear, concise, and easy to understand.
    You need to write a user manual that is accurate and up to date.
    You need to write a user manual that is comprehensive and covers all the features and capabilities of the platform.
    You need to include step by step guidance along with relevant screenshots from the platform to make sure that its easy to understand.
    You need to embed the screenshots in the text sections, to ensure flow of the content with the right images (like a textbook).
    """
)

verifier_agent = AgentDefinition(
    description="Expert Reviewer Agent who reviews the final work product. Makes sure that all the data is correct and accurate.",
    prompt="""
    You are an expert reviewer of technical writing.
    You are to go over all the work delivered by the Technical Writer Agent and review it for inaccuracies, hyperbole or other mistakes. 
    Specifically mistakes like: 
    * Wrong steps and information provided
    * Wrong screenshot mapped against the wrong section of documentation - This is very important and you need to check the screenshot against the content its been mapped against
    * Inaccurate information or hyperbole in the markdown documentation
    * Broken links like images and mermaid diagram images and fix

    You need to fix all of these before completing your job
    """
)


async def create_user_manual():
    instructions = """
    You are an experienced technical documentation expert. 
    You are tasked with creating a user manual for a product called OWASP Juice Shop. 
    The high-level details of the product are in the INPUT.md file. This file additionally has information about the URL, the credentials that you should use to log in to the application. 

    You need to do the following: 
    1. You need to read the INPUT.md file and understand the product and its capabilities, at a high level
    2. Use the 'memory' mcp server to track a list of TODOs based on the specifications from the INPUT.md file and track and update the TODOs as you go through the steps. 
    3. You need to use the research-agent to do a complete walkthrough of the application like a regular user. Use the agent-browser skill that uses a browser to open and use the application
    4. Then you need to call on the technical-writer-agent to write the user manual based on the information captured by the research agent and the TODOs that you have tracked. 
    5. Finally, you need to call on the verifier-agent to review the work product and make sure that all the data is correct and accurate.

    Finally you need to write a detailed modular guide in the output/ directory in markdown format with all images, mermaid diagrams, etc.
    If you're using mermaid diagrams, convert them into png and then embed in the markdown instead of embedding mermaid diagrams as code in the markdown.

    Note: DO NOT include, content, sections or topics that are not there in the platform. Only include content that exists and that you can verify
    """

    async def create_user_man():
        yield {
            "type": "user",
            "message": {
                "role": "user",
                "content": "Create a detailed user manual for OWASP Juice Shop application"
            }
        }

    options = ClaudeAgentOptions(
        max_turns = 10000,
        permission_mode="bypassPermissions",
        system_prompt=instructions,
        setting_sources=["project"],
        allowed_tools = APPROVED_TOOLS,
        mcp_servers={
            "memory": {
                "command": "npx",
                "args": [
                    "-y",
                    "@modelcontextprotocol/server-memory"
                ]
            },
        },
        agents = {
            "research-agent": research_agent,
            "technical-writer-agent": technical_writer_agent,
            "verifier-agent": verifier_agent,
        },
        cwd=os.path.join(os.getcwd(), "output")
    )

    async with ClaudeSDKClient(options) as client:
        await client.query(create_user_man())

        async for message in client.receive_response():
            if isinstance(message, AssistantMessage):
                for block in message.content:
                    if isinstance(block, TextBlock):
                        logger.info(block.text)
                    if isinstance(block, ThinkingBlock):
                        logger.info(block.thinking)
                    # if isinstance(block, ToolUseBlock):
                    #     logger.info(f"using {block.name} tool")



if __name__ == "__main__":
    asyncio.run(create_user_manual())