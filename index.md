---
layout: default
title: Home
permalink: /
description: "Amedeo Andriollo — PhD candidate in Economics at the University of Warwick. Research in econometrics, time series, and financial econometrics."
---

# About me

I am a PhD candidate in Economics at the [University of Warwick](https://warwick.ac.uk/fac/soc/economics/).

My research lies at the intersection of econometrics, time series, and financial econometrics.

My main supervisor is [Eric Renault](https://warwick.ac.uk/fac/soc/economics/staff/emrrenault/#). My second main supervisor is [Cesare Robotti](https://www.cesarerobotti.com).

{% assign visible_items = site.data.latest | slice: 0, 4 %}

{% if visible_items.size > 0 %}
## Latest

<ul class="latest-list">
{% for item in visible_items %}
<li class="latest-item">
  <a href="{{ item.link }}">
    {% if item.badge %}<span class="latest-badge">{{ item.badge }}</span>{% endif %}
    <span class="latest-text">{{ item.title }}</span>
  </a>
</li>
{% endfor %}
</ul>
{% endif %}
