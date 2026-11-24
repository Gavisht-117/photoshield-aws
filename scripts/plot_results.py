"""Plots CloudWatch metrics from a load test: CPU, instance count and request rate.

Usage: python scripts/plot_results.py --minutes 40   (the last 40 minutes)
Writes docs/results/scaling.png.
"""
import argparse
import datetime as dt
import os

import boto3
import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402

ASG_NAME = "cloud-ha-web-asg"
ALB_NAME = "cloud-ha-alb"


def series(cw, namespace, metric, dims, stat, start, end):
    resp = cw.get_metric_data(
        MetricDataQueries=[{
            "Id": "m",
            "MetricStat": {
                "Metric": {"Namespace": namespace, "MetricName": metric,
                           "Dimensions": [{"Name": k, "Value": v} for k, v in dims.items()]},
                "Period": 60,
                "Stat": stat,
            },
        }],
        StartTime=start, EndTime=end, ScanBy="TimestampAscending",
    )
    data = resp["MetricDataResults"][0]
    return data["Timestamps"], data["Values"]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--minutes", type=int, default=40)
    parser.add_argument("--region", default="us-east-1")
    parser.add_argument("--out", default=os.path.join(os.path.dirname(__file__), "..", "docs", "results", "scaling.png"))
    args = parser.parse_args()

    end = dt.datetime.now(dt.timezone.utc)
    start = end - dt.timedelta(minutes=args.minutes)
    cw = boto3.client("cloudwatch", region_name=args.region)
    lb_arn = boto3.client("elbv2", region_name=args.region).describe_load_balancers(
        Names=[ALB_NAME])["LoadBalancers"][0]["LoadBalancerArn"]
    lb_dim = lb_arn.split(":loadbalancer/")[1]

    cpu = series(cw, "AWS/EC2", "CPUUtilization", {"AutoScalingGroupName": ASG_NAME}, "Average", start, end)
    in_service = series(cw, "AWS/AutoScaling", "GroupInServiceInstances", {"AutoScalingGroupName": ASG_NAME},
                        "Maximum", start, end)
    desired = series(cw, "AWS/AutoScaling", "GroupDesiredCapacity", {"AutoScalingGroupName": ASG_NAME},
                     "Maximum", start, end)
    requests = series(cw, "AWS/ApplicationELB", "RequestCount", {"LoadBalancer": lb_dim}, "Sum", start, end)

    fig, axes = plt.subplots(3, 1, figsize=(10, 8), sharex=True)
    axes[0].plot(*cpu, color="#d9480f", marker=".", label="Average CPU (Auto Scaling group)")
    axes[0].axhline(50, color="#555", linestyle="--", linewidth=1, label="Scaling target 50%")
    axes[0].set_ylabel("CPU %")
    axes[0].set_ylim(0, 105)
    axes[1].step(*desired, where="post", color="#5b6cff", linestyle="--", label="Desired capacity")
    axes[1].step(*in_service, where="post", color="#2b8a3e", linewidth=2, label="In-service instances")
    axes[1].set_ylabel("Instances")
    axes[1].set_ylim(0, 5)
    axes[2].bar(requests[0], requests[1], width=0.0006, color="#868e96", label="Requests per minute (ALB)")
    axes[2].set_ylabel("Requests / min")
    for ax in axes:
        ax.grid(alpha=0.3)
        ax.legend(loc="upper left", fontsize=9)
    axes[2].xaxis.set_major_formatter(mdates.DateFormatter("%H:%M", tz=dt.timezone.utc))
    axes[2].set_xlabel("Time (UTC)")
    fig.suptitle("Load test: CPU-driven auto scaling of the PhotoShield web tier", fontsize=13)
    fig.tight_layout()
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    fig.savefig(args.out, dpi=150)
    print("Saved", os.path.abspath(args.out))


if __name__ == "__main__":
    main()
