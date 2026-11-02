from crewai import Agent, Crew, Process, Task , LLM
from crewai.project import CrewBase, agent, crew, task
from crewai.agents.agent_builder.base_agent import BaseAgent
from .tools.custom_tool import available_tools


llm = LLM(
    model="ollama/gpt-oss:20b-cloud",  # Replace with your pulled model
    base_url="http://localhost:11434"
) 
@CrewBase
class CodeReview():
    """CodeReview crew"""

    agents: list[BaseAgent]
    tasks: list[Task]

    @agent
    def code_review(self) -> Agent:
        return Agent(
            config=self.agents_config['code_review'], # type: ignore[index]
            verbose=True,
            tools=available_tools,
            llm=llm
        )

    @agent
    def code_structure_info(self) -> Agent:
        return Agent(
            config=self.agents_config['code_structure_info'], # type: ignore[index]
            verbose=True,
            tools=available_tools,
            llm=llm
        )
    @agent
    def code_vulnerability_check(self) -> Agent:
        return Agent(
            config=self.agents_config['code_vulnerability_check'], # type: ignore[index]
            verbose=True,
            tools=available_tools,
            llm=llm
        )
    @agent
    def write_prompt(self) -> Agent:
        return Agent(
            config=self.agents_config['write_prompt'], # type: ignore[index]
            verbose=True,
            tools = available_tools,
            llm=llm
        )


    @task
    def code_review_task(self) -> Task:
        return Task(
            config=self.tasks_config['code_review_task'], # type: ignore[index]
        )

    @task
    def code_structure_info_task(self) -> Task:
        return Task(
            config=self.tasks_config['code_structure_info_task'], # type: ignore[index]
            output_file='structure.json'
        )
    @task
    def code_vulnerability_check_task(self) -> Task:
        return Task(
            config=self.tasks_config['code_vulnerability_check_task'], # type: ignore[index]
            output_file='vulnerabilities.json'
        )

    @task
    def write_prompt_task(self) -> Task:
        return Task(
            config=self.tasks_config['write_prompt_task'], # type: ignore[index]
            output_file='prompt.md'
        )




    @crew
    def crew(self) -> Crew:
        """Creates the CodeReview crew"""
        return Crew(
            agents=self.agents,
            tasks=self.tasks,
            process=Process.sequential,
            verbose=True,
        )
