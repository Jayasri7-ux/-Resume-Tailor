import pytest
from fastapi.testclient import TestClient
from backend import main

client = TestClient(main.app)


def test_strict_taxonomy_extraction_eliminates_all_irrelevant_words():
    jd = """
    We are seeking a passionate, proactive, and dynamic Senior Full Stack Engineer to join our fast-paced, 
    collaborative, cross-functional team in an innovative environment.
    Key Responsibilities:
    - Responsible for designing, building, and delivering high-quality, scalable web applications.
    - Collaborate with product managers, designers, and business stakeholders to ensure successful delivery.
    - Foster team culture, drive excellence, and mentor junior developers in best practices.
    Technical Requirements:
    - Strong proficiency in Python, Django, and React.
    - Hands-on experience with PostgreSQL, MongoDB, and Redis databases.
    - Experience building REST APIs and deploying with Docker and Git on AWS.
    - Solid foundation in Data Structures, OOP, and CI/CD pipelines.
    - Knowledge of Pandas or Data Analysis is a plus.
    """
    extracted = main.extract_keywords(jd)

    # 1. Real technical entities MUST be extracted
    expected_tech = {
        "python", "django", "react", "postgresql", "mongodb", "redis",
        "docker", "git", "aws", "data structures", "oop",
        "ci/cd", "pandas", "data analysis"
    }
    for tech in expected_tech:
        assert tech in extracted, f"Expected tech '{tech}' was not extracted!"
    assert ("rest api" in extracted or "rest apis" in extracted), "REST API / APIs was not extracted!"

    # 2. Irrelevant English words, verbs, and corporate fluff MUST NEVER be extracted
    forbidden_words = {
        "seeking", "passionate", "proactive", "dynamic", "senior", "join",
        "fast-paced", "collaborative", "cross-functional", "team", "innovative",
        "environment", "responsible", "designing", "building", "delivering",
        "high-quality", "scalable", "applications", "collaborate", "product",
        "managers", "designers", "ensure", "successful", "delivery",
        "foster", "culture", "drive", "excellence", "mentor", "developers",
        "practices", "strong", "proficiency", "hands-on", "experience",
        "experience", "solid", "foundation", "knowledge", "plus"
    }
    extracted_forbidden = extracted.intersection(forbidden_words)
    assert not extracted_forbidden, f"Irrelevant words were extracted: {extracted_forbidden}"


def test_core_domain_inference():
    # Full Stack JD
    fs_jd = "Build web applications using Python, Django, React, and MongoDB."
    fs_tech = main.extract_keywords(fs_jd)
    domain_info = main.detect_core_domain(fs_jd, fs_tech, "Full Stack Developer")
    assert domain_info["detected"] is True
    assert domain_info["domain"] == "Full Stack Development"

    # Backend JD
    be_jd = "Develop backend microservices and REST APIs using Python, Flask, FastAPI, and PostgreSQL."
    be_tech = main.extract_keywords(be_jd)
    domain_info_be = main.detect_core_domain(be_jd, be_tech, "Backend Developer")
    assert domain_info_be["detected"] is True
    assert domain_info_be["domain"] == "Backend Development"

    # Data Engineering JD
    data_jd = "Build automated ETL pipelines using Python, SQL, Pandas, NumPy, and Tableau dashboards."
    data_tech = main.extract_keywords(data_jd)
    domain_info_data = main.detect_core_domain(data_jd, data_tech, "Data Analyst")
    assert domain_info_data["detected"] is True
    assert domain_info_data["domain"] == "Data Engineering & Analytics"


def test_non_technical_jd_stops_and_raises_422():
    non_tech_jd = "Looking for a retail store associate to count cash, fold clothes, and assist shoppers."
    with pytest.raises(Exception) as exc_info:
        main.score_resume("John Doe\nExperience\nCashier", non_tech_jd)
    assert "No core technical domain found" in str(exc_info.value.detail) or "more detail" in str(exc_info.value.detail)


def test_missing_technologies_have_explanations_and_pending_status():
    resume_text = "Jaya Sri Palli\nSkills: Python, Django, React, SQL, MongoDB, Git"
    jd_text = "Developer needs Python, Django, React, Docker, Redis, and AWS."
    analysis = main.score_resume(resume_text, jd_text, "Full Stack Developer")

    assert "missing_technologies" in analysis
    missing = analysis["missing_technologies"]
    missing_tools = {m["tool"] for m in missing}

    # Docker, Redis, AWS are missing from the resume
    assert "docker" in missing_tools
    assert "redis" in missing_tools
    assert "aws" in missing_tools

    # Python, Django, React are present in the resume
    assert "python" not in missing_tools
    assert "django" not in missing_tools
    assert "react" not in missing_tools

    # Each missing technology must have an explanation and pending status
    for item in missing:
        assert item["status"] == "pending"
        assert item["explanation"]
        assert item["domain"]


def test_incompatible_role_warning():
    compat = main.match_tier1_role_eligibility("iOS Swift Developer", "Build native iPhone apps using Swift and Xcode.")
    assert compat["eligibility_status"] == "Incompatible Domain Warning"
    assert "warning" in compat
    assert "iOS" in compat["domain"]
