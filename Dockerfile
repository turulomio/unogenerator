FROM ubuntu
MAINTAINER turulomio
RUN apt update
RUN apt --yes install python3-pip
RUN apt --yes install libreoffice
RUN apt --yes install imagemagick
RUN apt --yes install net-tools
RUN pip install unogenerator
RUN apt clean
EXPOSE 2002
CMD unogenerator_start; watch -n 0.1 'echo hola'

