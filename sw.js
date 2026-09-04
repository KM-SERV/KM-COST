/* KM Services — Coût de revient
   Fonctionnement hors connexion.

   Ce fichier n'a pas de numéro de version à incrémenter : chaque ouverture
   en ligne va chercher la dernière version des fichiers et remplace ce qui
   est en cache. Vous pouvez remplacer index.html sur GitHub sans rien
   toucher ici. */

const CACHE = 'km-cout';

const SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './apple-touch-icon.png'
];

/* Installation : on remplit le cache en ignorant le cache HTTP du navigateur,
   pour ne jamais y déposer une version périmée. */
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll(SHELL.map(u => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting())
      .catch(() => self.skipWaiting())
  );
});

/* Activation : ménage des anciens caches datés (km-cout-v1, etc.). */
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(noms => Promise.all(noms.filter(n => n !== CACHE).map(n => caches.delete(n))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;

  // On laisse passer les appels vers Google (sauvegarde Drive, polices).
  if (req.method !== 'GET') return;
  if (new URL(req.url).origin !== self.location.origin) return;

  // Ouverture de la page : réseau d'abord, cache en secours.
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(new Request(req.url, { cache: 'reload', credentials: 'same-origin' }))
        .then(rep => {
          const copie = rep.clone();
          caches.open(CACHE).then(c => c.put('./index.html', copie)).catch(() => {});
          return rep;
        })
        .catch(() => caches.match('./index.html').then(r => r || caches.match(req)))
    );
    return;
  }

  /* Icônes, manifeste : réponse immédiate depuis le cache, rafraîchissement
     en arrière-plan. Une icône modifiée apparaît à l'ouverture suivante,
     sans intervention. */
  e.respondWith(
    caches.match(req).then(cachee => {
      const reseau = fetch(req).then(rep => {
        const copie = rep.clone();
        caches.open(CACHE).then(c => c.put(req, copie)).catch(() => {});
        return rep;
      }).catch(() => cachee);
      return cachee || reseau;
    })
  );
});

/* Bouton « Vider le cache » de la page Paramètres. */
self.addEventListener('message', e => {
  if (e.data && e.data.type === 'vider') {
    caches.keys().then(noms => Promise.all(noms.map(n => caches.delete(n))));
  }
});
