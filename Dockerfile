# This defines the image that runs the backend.
# Typically, you'd invoke a script that builds the binary then invokes Docker.
# Note that you need Docker installed and running.

# Without specifying platform, an exec docker error can occur when running in k8s
# after the image was built in an Apple Silican Mac

FROM --platform=linux/amd64 jetty:9.4.49-jdk11
MAINTAINER Me "markus@molnify.com"
ADD ./target/xlport-2.0beta-SNAPSHOT /var/lib/jetty/webapps/ROOT
