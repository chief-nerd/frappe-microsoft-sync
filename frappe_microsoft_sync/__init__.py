__version__ = "0.0.1"

_oauth_patched = False


def patch_oauth():
    global _oauth_patched
    if _oauth_patched:
        return
    _oauth_patched = True

    import frappe
    import frappe.utils.oauth

    original_get_info_via_oauth = frappe.utils.oauth.get_info_via_oauth

    def mimirio_get_info_via_oauth(provider, code, decoder=None, id_token=False):
        import json

        flow = frappe.utils.oauth.get_oauth2_flow(provider)
        oauth2_providers = frappe.utils.oauth.get_oauth2_providers()

        args = {
            "data": {
                "code": code,
                "redirect_uri": frappe.utils.oauth.get_redirect_uri(provider),
                "grant_type": "authorization_code",
            }
        }

        if decoder:
            args["decoder"] = decoder

        session = flow.get_auth_session(**args)

        # Capture refresh token from the token response
        parsed_access = json.loads(session.access_token_response.text)
        refresh_token = parsed_access.get("refresh_token")

        if id_token:
            import jwt

            token = parsed_access["id_token"]
            info = jwt.decode(
                token, flow.client_secret, options={"verify_signature": False}
            )
        else:
            api_endpoint = oauth2_providers[provider].get("api_endpoint")
            api_endpoint_args = oauth2_providers[provider].get("api_endpoint_args")
            info = session.get(api_endpoint, params=api_endpoint_args).json()

            # GitHub email fallback (mirrors upstream frappe behavior)
            if provider == "github" and not info.get("email"):
                emails = session.get("/user/emails", params=api_endpoint_args).json()
                email_dict = next(filter(lambda x: x.get("primary"), emails))
                info["email"] = email_dict.get("email")

        if not (info.get("email_verified") or frappe.utils.oauth.get_email(info)):
            frappe.throw(
                frappe._("Email not verified with {0}").format(provider.title())
            )

        # Store refresh token for Microsoft providers
        if provider in ("office_365", "microsoft", "Microsoft") and refresh_token:
            email = frappe.utils.oauth.get_email(info)
            if email:
                frappe.cache().set_value(
                    f"ms_refresh_token:{email.lower()}",
                    refresh_token,
                    expires_in_sec=600,
                )

        return info

    frappe.utils.oauth.get_info_via_oauth = mimirio_get_info_via_oauth

    # Hook into login_oauth_user to persist the cached refresh token
    original_login_oauth_user = frappe.utils.oauth.login_oauth_user

    def mimirio_login_oauth_user(
        data, *, provider=None, state=None, generate_login_token=False
    ):
        res = original_login_oauth_user(
            data,
            provider=provider,
            state=state,
            generate_login_token=generate_login_token,
        )

        if provider in ("office_365", "microsoft", "Microsoft"):
            email = frappe.utils.oauth.get_email(data)
            if email:
                email = email.lower()
                refresh_token = frappe.cache().get_value(f"ms_refresh_token:{email}")
                if refresh_token:
                    from frappe.utils.password import set_encrypted_password
                    from frappe_microsoft_sync.microsoft_graph import MicrosoftGraphClient

                    client = MicrosoftGraphClient(email)
                    set_encrypted_password(
                        "MS Sync Settings",
                        client.settings.name,
                        refresh_token,
                        "refresh_token",
                    )
                    frappe.cache().delete_value(f"ms_refresh_token:{email}")

        return res

    frappe.utils.oauth.login_oauth_user = mimirio_login_oauth_user
