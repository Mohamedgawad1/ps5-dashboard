import json
import sys

from playwright.sync_api import sync_playwright


def main():
    with sync_playwright() as pw:
        b = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        ctx = b.contexts[0]
        for p in ctx.pages:
            try:
                st = p.evaluate(
                    """() => ({
                        url: location.href,
                        loggedIn: (typeof dalc !== 'undefined' && dalc.user && dalc.user.loggedIn === true) || false,
                        products: (typeof switchboardConfiguration !== 'undefined' && switchboardConfiguration.Products)
                            ? switchboardConfiguration.Products.length : -1,
                        loginField: document.getElementById('txtUserName') !== null,
                        title: document.title,
                    })"""
                )
                print(json.dumps(st, ensure_ascii=False))
            except Exception as e:
                print("ERR page:", str(e)[:150])
        b.close()


if __name__ == "__main__":
    sys.exit(main())