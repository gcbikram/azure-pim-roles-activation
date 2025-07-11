This change is to roll back the migration of the following Wave 1 Sprint 2b servers from Azure, restoring them to their original on-premises environment:

- WIN7305 SIT - DIP Web Server 2 (10.228.130.92)
- WIN7306 SIT - DIP Web Server 1 (10.228.130.91)
- WIN7307 SIT - DIP App Server 2 (10.228.130.98)
- WIN7308 SIT - DIP App Server 1 (10.228.130.99)
- WIN7309 SIT - BRE Server 1 (10.228.130.172)
- WIN7310 SIT - BRE Server 2 (10.228.130.173)
- WIN7313 SIT - Triage Server (10.228.130.121)

Rollback activities include:
1. Place SL1 & Nagios into maintenance mode
2. Update DNS to lower TTL
3. Shut down Azure servers
4. Power on on-premises servers (which have not been decommissioned)
5. Log in to each server with a local account and confirm functionality (including DC connectivity)
6. Force update server DNS
7. Update SL1 to use old IP addresses
8. Update application DNS entries
9. Remove SL1 from maintenance mode
10. Enable F5 health probes
11. Turn on F5 alerting
12. Address zScaler changes as required

- Revert F5 and Azure Load Balancer configurations based on dependency analysis

Servers will be restored in phases: Web servers first, then App, BRE, and finally Triage servers.

**Impact:**  
A full outage is required for all listed servers during the rollback, as connectivity between Azure and on-premises will be suspended. The outage is expected to last approximately one day to allow for migration, configuration, and testing.

*Note: OneNZ Zscaler changes will be managed and logged separately.*
