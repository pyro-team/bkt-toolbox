{% if site.google_analytics %}

(function() {
var ga = document.createElement('script'); ga.async = true;
ga.src = 'https://www.googletagmanager.com/gtag/js?id={{ site.google_analytics }}';
var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
})();

window.dataLayer = window.dataLayer || [];
function gtag(){dataLayer.push(arguments);}
gtag('js', new Date());
gtag('config', '{{ site.google_analytics }}');

{% endif %}