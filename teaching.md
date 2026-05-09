---
layout: default
title: Teaching
permalink: /teaching/
description: "Teaching experience and course materials by Giulio Rossetti at University of Warwick, including Financial Econometrics, Matrix Algebra, and Statistical Methods."
keywords: "teaching, university of warwick, financial econometrics, matrix algebra, statistical methods, course materials"
---

{% if site.data.teaching.current_courses and site.data.teaching.current_courses.size > 0 %}
# Current Teaching

{% for course in site.data.teaching.current_courses %}
<div class="course-card">
  <div class="course-header">
    <h2>{{ course.code }} - {{ course.name }}</h2>
    <span class="course-meta">{{ course.level }} | {{ course.term }}</span>
  </div>

  {% if course.description %}
  <p class="course-description">{{ course.description }}</p>
  {% endif %}

  <div class="weeks-container">
    {% for week in course.weeks %}
    <details class="week-accordion" {% if forloop.last %}open{% endif %}>
      <summary class="week-header">
        <span class="week-number">Week {{ week.week }}</span>
        <span class="week-toggle"></span>
      </summary>
      <div class="week-content">
        {% for topic in week.topics %}
        <div class="topic-item">
          <span class="topic-name">{{ topic.name }}</span>
          <div class="material-links">
            {% if topic.slides %}
            <a href="{{ site.baseurl }}/assets/slides/{{ topic.slides }}" target="_blank" class="material-btn slides-btn" onclick="trackDownload('{{ course.code }} - Week {{ week.week }} - {{ topic.name }} Slides')">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
              Slides
            </a>
            {% endif %}
            {% if topic.annotated %}
            <a href="{{ site.baseurl }}/assets/slides/{{ topic.annotated }}" target="_blank" class="material-btn annotated-btn" onclick="trackDownload('{{ course.code }} - Week {{ week.week }} - {{ topic.name }} Annotated')">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>
              Annotated
            </a>
            {% endif %}
            {% if topic.extras %}
              {% for extra in topic.extras %}
              <a href="{{ site.baseurl }}/assets/slides/{{ extra.file }}" target="_blank" class="material-btn extra-btn" onclick="trackDownload('{{ course.code }} - Week {{ week.week }} - {{ extra.name }}')">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                {{ extra.name }}
              </a>
              {% endfor %}
            {% endif %}
          </div>
        </div>
        {% endfor %}
      </div>
    </details>
    {% endfor %}
  </div>
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
      <span class="course-code">{{ course.code }}</span>
      <span class="course-name">{{ course.name }}</span>
      <span class="course-level">{{ course.level }}</span>
    </div>
    {% endfor %}
  </div>
</div>
{% endfor %}

{% if older_years.size > 0 %}
<details class="older-teaching-accordion">
  <summary class="older-teaching-header">
    <span>Older (2021-2023)</span>
    <span class="older-toggle"></span>
  </summary>
  <div class="older-teaching-content">
    {% for year_data in older_years %}
    <div class="year-section">
      <h2 class="year-heading">{{ year_data.year }}</h2>
      <div class="courses-grid">
        {% for course in year_data.courses %}
        <div class="past-course-card">
          <span class="course-code">{{ course.code }}</span>
          <span class="course-name">{{ course.name }}</span>
          <span class="course-level">{{ course.level }}</span>
        </div>
        {% endfor %}
      </div>
    </div>
    {% endfor %}
  </div>
</details>
{% endif %}
