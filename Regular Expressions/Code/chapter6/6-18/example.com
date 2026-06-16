$TTL 86400
@	IN	SOA	host.example.com.	root@example.com	(
				17 ; serial
				28800 ; refresh
				14400 ; retry
				3600000 ; expire
				86400 ; ttl
				)



	IN	NS	ns1.example.com.

host1          IN A 10.0.0.1
host2    IN CNAME    host1.example.com.
