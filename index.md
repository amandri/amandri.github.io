---
layout: default
title: Home
permalink: /
description: "Amedeo Andriollo — "
---

# About me

I am an Assistant Professor of Finance at the [Université Paris Dauphine – PSL (DRM)](https://dauphine.psl.eu). 

My research lies at the intersection of econometrics, time series, and financial econometrics.

I graduated from the [University of Warwick](https://warwick.ac.uk/fac/soc/economics/) under the supervisor of [Eric Renault](https://warwick.ac.uk/fac/soc/economics/staff/emrrenault/#), [Cesare Robotti](https://www.cesarerobotti.com), and [Giovanni Ricco](https://www.giovanni-ricco.com)

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
