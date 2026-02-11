./build_and_upload.sh
kubectl apply -f kube-xlport.yaml
echo "Stopping old version"
kubectl scale deployment/xlport --replicas=0
echo "Starting new version"
kubectl scale deployment/xlport --replicas=1
