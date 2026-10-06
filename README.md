# EwsReadiness

Diagnostic en lecture seule pour préparer la fin d’**Exchange Web Services (EWS)** dans Exchange Online.

Depuis le **10 octobre 2026**, un locataire configuré avec `EWSEnabled = True` doit aussi avoir une liste d’autorisation d’applications (`EWSAllowedAppIDs`). Sans elle, tout le trafic EWS est bloqué. Les locataires restés à la valeur par défaut verront ensuite `EWSEnabled` passer à `False` et EWS sera désactivé pour de bon le **1er avril 2027**.

EwsReadiness répond à une seule question : **quelles applications vont cesser de fonctionner si rien n’est fait ?**

📝 Article complet : [Exchange Online : la fin d’EWS passe à la vitesse supérieure le 10 octobre](https://blog.trivium365.com/exchange-online-fin-ews-liste-autorisation-10-octobre-2026/)

## Ce que fait l’outil

1. Lit `EwsEnabled` et la liste `EWSAllowedAppIDs` du locataire.
2. Lit l’export CSV du rapport **Utilisation d’EWS** (facultatif mais recommandé).
3. Repère les applications Entra qui détiennent une permission EWS (`full_access_as_app` ou `EWS.AccessAsUser.All`).
4. Traduit chaque AppID en nom d’application et classe le résultat :

| Statut | Signification |
|---|---|
| **À AJOUTER** | Utilisée mais absente de la liste : elle sera bloquée |
| **À VALIDER** | Autorisée mais jamais vue dans le rapport d’utilisation |
| **PERMISSION SEULEMENT** | Détient une permission EWS sans usage ni autorisation |
| **OK** | Utilisée et autorisée |

Il produit un tableau à l’écran, un fichier CSV et un rapport HTML.

**L’outil ne modifie rien.** Aucune commande `Set-` n’est exécutée.

## Prérequis

- PowerShell 7 ou Windows PowerShell 5.1
- Modules : `ExchangeOnlineManagement`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Identity.SignIns`
- Un rôle Exchange en lecture (par exemple Lecteur général)
- Les autorisations Microsoft Graph `Application.Read.All` et `Directory.Read.All` (consentement demandé à la connexion)

```powershell
Install-Module ExchangeOnlineManagement, Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns -Scope CurrentUser
```

## Utilisation

1. Exportez le rapport d’utilisation : Centre d’administration Microsoft 365 > Rapports > Utilisation > Exchange > onglet **Utilisation d’EWS**. Choisissez 90 jours puis exportez en CSV.
2. Lancez l’outil :

```powershell
.\EwsReadiness.ps1 -UsageReportCsv .\EWSUsage.csv -OutputFolder .\rapports
```

Déjà connecté à Exchange Online et à Microsoft Graph ? Ajoutez `-SkipConnect`.

## Exemple de résultat

```
Statut               Application              AppId                                 Autorisee Utilisee
------               -----------              -----                                 --------- --------
À AJOUTER            Sauvegarde Contoso       3f2a…                                 False     True
À VALIDER            Ancien connecteur CRM    9b41…                                 True      False
OK                   Microsoft Office         d3590ed6-52b3-4102-aeff-aad2292ab01c  True      True
```

## Et ensuite ?

Une fois la liste validée, la modification se fait avec `Set-OrganizationConfig -EwsAllowedAppIDs "AppID1,AppID2"`. Cette commande **remplace** toute la liste : relisez d’abord la liste actuelle avec `Get-OrganizationConfig -RetrieveEwsOperationAccessPolicy`, ajoutez le nouvel AppID puis réécrivez l’ensemble. Le changement peut prendre jusqu’à 24 heures.

## Limites connues

- Le rapport d’utilisation est agrégé chaque semaine, avec jusqu’à 10 jours de décalage. Une application lancée rarement peut ne pas y figurer.
- Sans le CSV, l’outil ne peut pas savoir quelles applications utilisent réellement EWS.
- Seuls les nuages commerciaux sont visés par le calendrier du 10 octobre. Les nuages gouvernementaux suivent un calendrier distinct.

## Sources

- [EWS Deprecation Is Here : What This Means To You](https://techcommunity.microsoft.com/blog/exchange/ews-deprecation-is-here-%E2%80%93-what-this-means-to-you/4561431) (Exchange Team)
- [Introducing EWSAllowedAppIDs](https://techcommunity.microsoft.com/blog/exchange/introducing-ewsallowedappids-preparing-for-the-final-phase-of-ews-retirement/4529471) (Exchange Team)
- [Deprecation of Exchange Web Services in Exchange Online](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online) (Microsoft Learn)
- [EWS Usage Report](https://learn.microsoft.com/en-us/microsoft-365/admin/activity-reports/ews-usage) (Microsoft Learn)

## Licence

MIT. Publié par Charles Jenkins, [blog.trivium365.com](https://blog.trivium365.com).
