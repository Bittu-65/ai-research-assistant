from crewai import Crew

from agents.research_specialist import research_specialist_agent
from agents.data_analyst import data_analyst_agent
from agents.content_writer import content_writer_agent
from agents.ppt_maker import ppt_maker_agent
from agents.gap_identifier import gap_identifier_agent
from agents.literature_reviewer import literature_reviewer_agent
from agents.paper_comparator import paper_comparator_agent
from agents.citation_generator import citation_reference_generator_agent
from agents.problem_statement_generator import problem_statement_generator_agent

from tasks.research_task import research_task
from tasks.analysis_task import analysis_task
from tasks.writing_task import writing_task
from tasks.ppt_task import ppt_task
from tasks.gap_identification_task import gap_identification_task
from tasks.literature_review_task import literature_review_task
from tasks.paper_comparison_task import paper_comparison_task
from tasks.citation_generation_task import citation_reference_task
from tasks.problem_statement_generation_task import problem_statement_task


# ── Existing crews (unchanged) ────────────────────────────────────────────────

research_crew = Crew(
    agents=[
        research_specialist_agent,
        data_analyst_agent,
        gap_identifier_agent,
        content_writer_agent,
    ],
    tasks=[
        research_task,
        analysis_task,
        gap_identification_task,
        writing_task,
    ],
    verbose=True
)

research_crew_with_ppt = Crew(
    agents=[
        research_specialist_agent,
        data_analyst_agent,
        gap_identifier_agent,
        content_writer_agent,
        ppt_maker_agent,
    ],
    tasks=[
        research_task,
        analysis_task,
        gap_identification_task,
        writing_task,
        ppt_task,
    ],
    verbose=True
)

literature_review_crew = Crew(
    agents=[
        research_specialist_agent,
        data_analyst_agent,
        gap_identifier_agent,
        literature_reviewer_agent,
    ],
    tasks=[
        research_task,
        analysis_task,
        gap_identification_task,
        literature_review_task,
    ],
    verbose=True
)

paper_comparison_crew = Crew(
    agents=[
        paper_comparator_agent,
    ],
    tasks=[
        paper_comparison_task,
    ],
    verbose=True
)

# ── New: Citation & Reference Generator crew ──────────────────────────────────
# Runs the full research + analysis + writing pipeline first, then the citation
# agent extracts all sources and produces a formatted APA reference list with
# an inline citation guide.
citation_reference_crew = Crew(
    agents=[
        research_specialist_agent,
        data_analyst_agent,
        gap_identifier_agent,
        content_writer_agent,
        citation_reference_generator_agent,
    ],
    tasks=[
        research_task,
        analysis_task,
        gap_identification_task,
        writing_task,
        citation_reference_task,
    ],
    verbose=True
)

# ── New: Problem Statement Generator crew ─────────────────────────────────────
# Runs research, analysis, and gap identification, then the problem statement
# agent synthesises the findings into a structured academic problem statement
# with research questions, objectives, and scope definition.
problem_statement_crew = Crew(
    agents=[
        research_specialist_agent,
        data_analyst_agent,
        gap_identifier_agent,
        problem_statement_generator_agent,
    ],
    tasks=[
        research_task,
        analysis_task,
        gap_identification_task,
        problem_statement_task,
    ],
    verbose=True
)
