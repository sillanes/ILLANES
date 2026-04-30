 DomReady.ready(function() {
    var companyCustomLogo = document.getElementById('custom-logo'),
        companyParentCustomLogo = document.getElementById('custom-parent-logo');
    companyCustomLogo.addEventListener('load', function(event) {
            companyLogoSuccess = true;
            hide(companyCustomLogo);
            show(event.target, document.getElementById('custom-logo-spinner'));
        });
    companyCustomLogo.addEventListener('error', function(event) {
            if (!companyLogoSuccess) {
                show(document.getElementById('default-logo'), document.getElementById('custom-logo-spinner'));
            }
        });
    companyCustomLogo.src = companyLogoFolder + companyLogoImg;

    companyParentCustomLogo.addEventListener('load', function(event) {
            parentCompanyLogoSuccess = top.alternateFooter = true;
            hide(companyParentCustomLogo);
            show(event.target, document.getElementById('custom-parent-logo-spinner'));
        });
    companyParentCustomLogo.addEventListener('error', function(event) {
            if (!parentCompanyLogoSuccess) {
                show(document.getElementById('default-parent-logo'), document.getElementById('custom-parent-logo-spinner'));
            }
        });
    companyParentCustomLogo.src = parentCompanyLogoFolder + companyLogoImg;
});

function show(image, spinner) {
    image.style.display = 'inline-block';
    if (spinner) {
        spinner.style.display = 'none';
    }
}

function hide(image, spinner) {
    image.style.display = 'none';
    if (spinner) {
        spinner.style.display = 'inline-block';
    }
}

