---
layout: default
title: Teaching
permalink: /teaching/
description: "Teaching experience and course materials by Amedeo Andriollo at Dauphine University, including Artificial Intelligence for Corporate Finance, Asset Pricing Theory, and Empirical Asset Pricing."
keywords: "teaching, dauphine university, artificial intelligence, corporate finance, asset pricing, course materials"
---

{% if site.data.teaching.current_courses and site.data.teaching.current_courses.size > 0 %}
# Current Teaching

{% for course in site.data.teaching.current_courses %}
<div class="course-card">
  <div class="course-header">
    <h2>{% if course.code %}{{ course.code }} - {% endif %}{{ course.name }}</h2>
    {% if course.level or course.term %}
    <span class="course-meta">
      {% if course.level %}{{ course.level }}{% endif %}{% if course.level and course.term %} | {% endif %}{% if course.term %}{{ course.term }}{% endif %}
    </span>
    {% endif %}
  </div>

  {% if course.description %}
  <p class="course-description">{{ course.description }}</p>
  {% endif %}

  {% if course.weeks %}
  <div class="weeks-container">
    {% for week in course.weeks %}
    <details class="week-accordion" {% if forloop.first %}open{% endif %}>
      <summary class="week-header">
        <span class="week-number">Week {{ week.week }}</span>
        <span class="week-toggle"></span>
      </summary>
      <div class="week-content">
        {% for topic in week.topics %}
        <div class="topic-item">
          <span class="topic-name">{{ topic.name }}</span>
          <div class="material-links">

            {% if topic.materials %}
              {% for material in topic.materials %}
              <a href="{{ site.baseurl }}{{ material.path }}" target="_blank" class="material-btn extra-btn" onclick="trackDownload('{{ course.name }} - Week {{ week.week }} - {{ material.name }}')">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                {{ material.name }}
              </a>
              {% endfor %}
            {% endif %}

            {% if topic.material_folder %}
              {% assign pdf_count = 0 %}
              {% for file in site.static_files %}
                {% if file.path contains topic.material_folder and file.extname == '.pdf' %}
                  {% assign pdf_count = pdf_count | plus: 1 %}
                  <a href="{{ site.baseurl }}{{ file.path }}" target="_blank" class="material-btn slides-btn" onclick="trackDownload('{{ course.name }} - Week {{ week.week }} - {{ file.name }}')">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
                    {{ file.name | remove: '.pdf' | replace: '_', ' ' }}
                  </a>
                {% endif %}
              {% endfor %}
            {% endif %}

          </div>
        </div>
        {% endfor %}
      </div>
    </details>
    {% endfor %}
  </div>
  {% endif %}
</div>
{% endfor %}

<hr class="section-divider">
{% endif %}

# Past Teaching

{% assign recent_years = site.data.teaching.past_teaching | where_exp: "item", "item.year >= 2024" %}
{% assign older_years = site.data.teaching.past_teaching | where_exp: "item", "item.year < 2024" %}

{% for year_data in recent_years %}
<div class="year-section">
  <h2 class="year-heading">{{ year_data.year }}</h2>
  <div class="courses-grid">
    {% for course in year_data.courses %}
    <div class="past-course-card">
      {% if course.code %}<span class="course-code">{{ course.code }}</span>{% endif %}
      <span class="course-name">{{ course.name }}</span>
      {% if course.level %}<span class="course-level">{{ course.level }}</span>{% endif %}
      {% if course.term %}<span class="course-level">{{ course.term }}</span>{% endif %}

      {% if course.weeks %}
      <details class="week-accordion">
        <summary class="week-header">
          <span class="week-number">Course materials</span>
          <span class="week-toggle"></span>
        </summary>
        <div class="week-content">
          {% for week in course.weeks %}
            {% for topic in week.topics %}
            <div class="topic-item">
              <span class="topic-name">Week {{ week.week }} — {{ topic.name }}</span>
              <div class="material-links">
                {% if topic.materials %}
                  {% for material in topic.materials %}
                  <a href="{{ site.baseurl }}{{ material.path }}" target="_blank" class="material-btn extra-btn" onclick="trackDownload('{{ course.name }} - Week {{ week.week }} - {{ material.name }}')">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                    {{ material.name }}
                  </a>
                  {% endfor %}
                {% endif %}
              </div>
            </div>
            {% endfor %}
          {% endfor %}
        </div>
      </details>
      {% endif %}
    </div>
    {% endfor %}
  </div>
</div>
{% endfor %}

{% if older_years.size > 0 %}
<details class="older-teaching-accordion">
  <summary class="older-teaching-header">
    <span>Older teaching</span>
    <span class="older-toggle"></span>
  </summary>
  <div class="older-teaching-content">
    {% for year_data in older_years %}
    <div class="year-section">
      <h2 class="year-heading">{{ year_data.year }}</h2>
      <div class="courses-grid">
        {% for course in year_data.courses %}
        <div class="past-course-card">
          {% if course.code %}<span class="course-code">{{ course.code }}</span>{% endif %}
          <span class="course-name">{{ course.name }}</span>
          {% if course.level %}<span class="course-level">{{ course.level }}</span>{% endif %}
          {% if course.term %}<span class="course-level">{{ course.term }}</span>{% endif %}
        </div>
        {% endfor %}
      </div>
    </div>
    {% endfor %}
  </div>
</details>
{% endif %}
