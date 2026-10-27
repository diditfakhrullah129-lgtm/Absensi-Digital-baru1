(function () {
  if (window.google && window.google.script && window.google.script.run) return; // di dalam Apps Script

  var API_URL = window.GAS_API_URL;
  var memToken = null; // cadangan bila localStorage diblokir (iframe Canva)

  function getToken() {
    try {
      var s = JSON.parse(localStorage.getItem('absensiAppSession') || 'null');
      if (s && s.token) return s.token;
    } catch (e) {}
    return memToken;
  }

  function call(fn, args) {
    return fetch(API_URL, {
      method: 'POST',
      // text/plain sengaja dipakai agar tidak memicu preflight CORS (Apps Script tidak menanganinya)
      body: JSON.stringify({ fn: fn, args: args, token: getToken() })
    })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) throw new Error(j.error || 'Server error');
        if (fn === 'login' && j.result && j.result.success) memToken = j.result.token;
        if (fn === 'logoutUser') memToken = null;
        return j.result;
      });
  }

  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (f) { return runner(f, fail); };
        if (name === 'withFailureHandler') return function (f) { return runner(ok, f); };
        if (name === 'withUserObject') return function () { return runner(ok, fail); };
        return function () {
          var args = Array.prototype.slice.call(arguments);
          call(String(name), args).then(
            function (res) { if (ok) ok(res); },
            function (err) { if (fail) fail(err); else console.error(name, err); }
          );
        };
      }
    });
  }

  window.google = window.google || {};
  window.google.script = { run: runner(null, null), host: { close: function () {} } };
})();
