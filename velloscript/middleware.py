class SecurityHeadersMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)
        
        # CSP Normal (Mengizinkan inline styles & scripts agar web berfungsi normal)
        response['Content-Security-Policy'] = "default-src 'self'; style-src 'self' 'unsafe-inline' https://cdnjs.cloudflare.com https://fonts.googleapis.com https://cdn.jsdelivr.net; script-src 'self' 'unsafe-inline' 'unsafe-eval' https://cdnjs.cloudflare.com https://code.jquery.com https://cdn.jsdelivr.net; img-src 'self' data: https://ui-avatars.com; font-src 'self' https://cdnjs.cloudflare.com https://fonts.gstatic.com; frame-ancestors 'none'; form-action 'self';"
        
        # Strict Transport Security (HSTS)
        response['Strict-Transport-Security'] = "max-age=31536000; includeSubDomains; preload"
        
        # Anti Sniffing
        response['X-Content-Type-Options'] = "nosniff"
        
        # Anti Clickjacking
        response['X-Frame-Options'] = "DENY"

        # Cross-Domain Misconfiguration Fix
        response['Access-Control-Allow-Origin'] = "https://velloscript.pythonanywhere.com"

        return response
