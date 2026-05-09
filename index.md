---
layout: default
title: Home
permalink: /
description: "Amedeo Andriollo — AP of Finance at Université Paris Dauphine – PSL (DRM) in econometrics and financial econometrics, with a focus on empirical asset pricing."
---

# About Amedeo

I am an Assistant Professor of Finance at the [Université Paris Dauphine – PSL (DRM)](https://dauphine.psl.eu). 

My research lies at the intersection of econometrics, time series, and financial econometrics.

I graduated from the [University of Warwick](https://warwick.ac.uk/fac/soc/economics/) under the supervisor of [Eric Renault](https://warwick.ac.uk/fac/soc/economics/staff/emrrenault/#), [Cesare Robotti](https://www.cesarerobotti.com), and [Giovanni Ricco](https://www.giovanni-ricco.com)

<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Person",
  "name": "Amedeo Andriollo",
  "givenName": "Amedeo",
  "familyName": "Andriollo",
  "jobTitle": "Assistant Professor",
  "affiliation": {
    "@type": "CollegeOrUniversity",
    "name": "Université Paris Dauphine – PSL (DRM)",
    "url": "https://dauphine.psl.eu"
  },
  "alumniOf": {
    "@type": "CollegeOrUniversity",
    "name": "University of Warwick — Warwick Business School"
  },
  "knowsAbout": [
    "Econometrics",
    "Financial Econometrics",
    "Asset Pricing",
    "Time-series Econometrics",
    "Factor Models"
  ],
  "url": "https://amandri.github.io",
  "image": "https://giuliorossetti94.github.io/assets/img/ame2.png",
  "sameAs": [
    "https://github.com/amandri",
    "https://fr.linkedin.com/in/amedeo-andriollo-7a531a133/en",
    "https://x.com/amedandr"
  ]
}
</script>

{% assign visible_items = site.data.latest | slice: 0, 4 %}

{% if visible_items.size > 0 %}
## Latest

<ul class="latest-list">
{% for item in visible_items %}
<li class="latest-item">
  <a href="{{ item.link }}" onclick="gtag('event', 'latest_click', { event_category: 'Engagement', event_label: '{{ item.title | escape }}' });">
    {% if item.badge %}<span class="latest-badge">{{ item.badge }}</span>{% endif %}
    <span class="latest-text">{{ item.title }}</span>
  </a>
</li>
{% endfor %}
</ul>
{% endif %}
